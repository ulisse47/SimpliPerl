#!/usr/bin/perl
use strict;
use warnings;
use DBI;
use File::Basename;
use File::Copy;
use File::Path;

my $sourcePath = "E:\\FileServer\\Ricevute";
# monto cartella web Y: su http://mail.unioneartigiani.it/repository/
my $destPath = "Y:";
my $logFilePath = "$sourcePath\\Log\\Trasf.log";

# DATA SOURCE utenti
my $usersDatabaseServer = "webserver";
my $usersDatabaseName = "intranet";
my $usersDatabaseUser = "root";
my $usersDatabasePassword = "orione";

# DATA SOURCE documenti
my $docsDatabaseServer = "mail.unioneartigiani.it";
my $docsDatabaseName = "unioneartigiani";
my $docsDatabaseUser = "unioneartigiani";
my $docsDatabasePassword = "unioneartigiani";

sub getTimestamp() {
    my ($sec, $min, $hour, $mday, $mon, $year, $wday, $yday, $isdst) = localtime(time);
    $year += 1900;
    $mon += 1;
    return join('/', $mday, $mon, $year) . " " .  join(':', $hour, $min, $sec);
}

my $usersDbh = DBI->connect("DBI:mysql:database=" . $usersDatabaseName . ";host=". $usersDatabaseServer,
    $usersDatabaseUser, $usersDatabasePassword, {
      RaiseError => 0, ### if =0 Don't report errors via die
      PrintError => 0  ### if =0 Don't report errors via warn
    }
) or die "Unable to connect to MySQL users database: $DBI::errstr";
my $usersByCodPagSth = $usersDbh->prepare("SELECT CF, RagSoc, CodPag, Ufficio FROM ArchivioUtenti WHERE CodPag = ?")
    or die "Couldn't prepare statement, aborting";
my $usersByCFSth = $usersDbh->prepare("SELECT CF, RagSoc, CodPag, Ufficio FROM ArchivioUtenti WHERE CF = ?")
    or die "Couldn't prepare statement, aborting";

my $docsDbh = DBI->connect("DBI:Pg:dbname=" . $docsDatabaseName . ";host=". $docsDatabaseServer,
    $docsDatabaseUser, $docsDatabasePassword, {
      RaiseError => 0, ### if =0 Don't report errors via die
      PrintError => 0  ### if =0 Don't report errors via warn
    }
) or die "Unable to connect to PostgreSQL docs database: $DBI::errstr";
my $docsByPathSth = $docsDbh->prepare("SELECT Path FROM ArkDoc WHERE Path = ?")
    or die "Couldn't prepare statement, aborting";
my $insertDocByPathStx = $docsDbh->prepare("INSERT INTO ArkDoc"
    . "(CF, Tipo, Doc, Data, Anno, Mese, Ufficio, Servizio, Classe, File, Pathfile, Pag, Spedito, Path)"
    . "VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)")
    or die "Couldn't prepare statement, aborting";
my $updateDocByPathStx = $docsDbh->prepare("UPDATE ArkDoc"
    . " SET CF=?, Tipo=?, Doc=?, Data=?, Anno=?, Mese=?, Ufficio=?, Servizio=?,"
    . " Classe=?, File=?, Pathfile=?, Pag=?, Spedito=?"
    . " WHERE Path = ?")
    or die "Couldn't prepare statement, aborting";

open (LISTA, ">" . $logFilePath);

print LISTA "+++ " . getTimestamp() . " Inizio elaborazione +++\n*** Lista segnalazioni\n";

my @suffixes  = (".pdf", ".txt", ".eml");
my @fileTypes = ("DM10", "MENS", "ArkPaghe");

# cerco file con naming SOURCE\fileType\year\office\month\file
for my $fileType ( @fileTypes ) {
    #print "$fileType -->  Esiste\n";
    print "$sourcePath\\$fileType", "\n";
    
    for my $yearDir ( <$sourcePath\\$fileType\\*> ) {
        my $year = (fileparse($yearDir))[0];
        printf " ^%s;\n", $year ;
        next if $year ne "2009";

        for my $officeDir ( <$yearDir\\*> ) {
            my $codUff = (fileparse($officeDir))[0];
            printf "\n  >%s", $codUff;
            # next if $codUff ne "0000";

            for my $monthDir ( <$officeDir\\*> ) {
                my $mese = (fileparse($monthDir))[0];
                printf "\n   >>%s", $mese;

                my @destDirs = <$monthDir\\*>;
                for my $destDir ( @destDirs ) {
                    my $dest = (fileparse($destDir))[0];
                    #printf "\t->Inizio Dest -%s;%s\n", $dest, scalar(@destDirs);

                    my @docPaths;
                    # controllo special case per codice ufficio 0000
                    if ( $codUff eq "0000" ) {
                        @docPaths = ( $dest );
                    } else {
                        @docPaths = <$destDir\\*.*>;
                    }
                    
                    for my $docPath ( @docPaths ) {                                     
                        my ( $doc, $directories, $suffix ) = fileparse($docPath, @suffixes);
                        next if $doc eq "_tutti";
                        #printf "\t\t->Inizio Doc +>%s;%s\n", $doc, scalar(@docPaths);

                        my ( $data, $servizio, $classe, $destDoc, $key ) = ( "", "", "", "", "" );
                        my $pag = 1;
                        my $selKey = "CF";                      

                        printf "\$fileType=%s, \$dest=%s, \$doc=%s, \$destDoc=%s\n", $fileType, $dest, $doc, $destDoc;
                                  
                        for ( $fileType ) {
                            if ( $fileType =~ /DM10$/ or $fileType =~ /MENS$/ ) {
                                printf "\ncaso 1\n";
                                $data = "$year/$mese/01";
                                $servizio = "PG";
                                $classe = "InpsM" ; # InpsM=Mensile Inps
                                $destDoc = "$fileType$suffix";
                                $key = $doc;
                                $selKey = "CodPG";
                            } elsif ( $fileType =~ /ArkPaghe/ ) {
                                printf "\ncaso 2\n";
                                $data ="$year/$mese/01";
                                $servizio = "PG";
                                $classe = "Elab"; # Elaborati mensili
                                $destDoc = "$doc$suffix";
                                $key = $dest;
                                $selKey = "CodPG";
                            }
                        }

                        # controllo special case per codice ufficio 0000
                        if ( $codUff eq "0000" ) {
                            $selKey = "CF";
                        }
                        #print "\t\t\tIn elaborazione-->", $key, "\n";

                        printf "------> \$selKey=%s, \$codUff=%s, \$servizio=%s, \$destDoc=%s, \$key=%s\n",
                            $selKey, $codUff, $servizio, $destDoc, $key;

                        my ($rf, $rb, $re);
                        my ($ragSoc, $codPag, $ufficio, $cf);

                        if ( $selKey eq "CF" ) {
                            $key = substr("     " . $key, -16, 16);
                            if ($re = $usersByCFSth->execute($key) ) {
                                $rb = $usersByCFSth->bind_columns(\$cf, \$ragSoc, \$codPag, \$ufficio);
                                if ( $rf = $usersByCFSth->fetch ) {
                                    printf "+++ \t\t\t%s;%s;%s;%s;%s;%s;%s\n", $key, $ragSoc, $codPag,
                                        $ufficio, $key, $cf, $destDoc;
                                } else {
                                    print "+++ \t\t\tCF non esistente $key\n";
                                }
                            }
                            $docPath = $destDir;
                        } elsif ( $rf = $usersByCodPagSth->execute($key) ) {
                            my $rv = $usersByCodPagSth->bind_columns(\$cf, \$ragSoc, \$codPag, \$ufficio);

                            if ( $rf = $usersByCodPagSth->fetch ) {
                                #printf  ("\t\t\t%s;%s;%s;%-30s\n", $doc, $codPag, $cf, $ragSoc);
                            } else {
                                printf LISTA "*** Azienda inesistente: $doc\n";
                            }
                        } else {
                            print "***Errore database - $doc\n";
                        }

                        # INSERT
                        $cf =~ s/ +//;
                        
                        my $file = $destDoc;
                        my $filePath = "$destPath\\$year\\$ufficio\\$cf\\$mese";
                        my $path = $filePath . "\\$destDoc";

                        printf "\$selKey=%s, \$ragSoc=%s, \$codPag=%s, \$codUff=%s, \$ufficio=%s, \$servizio=%s,"
                            . " \$destDoc=%s, \$key=%s, \$cf=%s, \$destDoc=%s\n", $selKey, $ragSoc, $codPag,
                            $codUff, $ufficio, $servizio, $destDoc, $key, $cf, $destDoc;

                        my $stmt = $insertDocByPathStx;

                        # SE IL RECORD ESISTE USO STMT DI UPDATE
                        if ($docsByPathSth->execute($path) && $docsByPathSth->fetch) {
                            printf "record con path esistente, effettuo update\n";
                            $stmt = $updateDocByPathStx;
                        }

                        my $ri = $stmt->execute( $cf, $fileType, $destDoc, $data, $year, $mese,
                            $ufficio, $servizio, $classe, $file, $filePath, $pag, "n", $path )
                            or die "Can't execute SQL statement: $DBI::errstr\n";

                        printf "\$cf=%s, \$fileType=%s, \$path=%s, \$destDoc=%s, \$data=%s, \$year=%s, \$mese=%s,"
                            . " \$ufficio=%s, \$servizio=%s, \$classe=%s, \$file=%s, \$filePath=%s, \$pag=%s,"
                            . " \$dest=%s\n", $cf, $fileType, $path, $destDoc, $data, $year, $mese, $ufficio,
                            $servizio, $classe, $file, $filePath, $pag, $dest;
                        
                        if ($ri) {
                            #printf  "*** Documento inserito: $path-$destDir\n";
                            print "+";
                        } else {
                            #printf  "*** Mancano dati obbligatori: $path\n";
                            print "-";
                        }

                        my $directory = (fileparse($path))[1];
                        stat $directory or mkpath($directory, 0, 0777);

                        printf "\ntrasferisci: %s\t-->\t%s\n", $docPath, $path;

                        copy($docPath, $path) or die "Copy failed: $!";

                        printf LISTA ("%s\t-->\t%s\n", $docPath, $path);
                        #printf ("%s;%s;%s;%s;%s;%s;%s\n",$year, $ufficio, $cf, $mese, $docPath, $dest, $path);
                        #print "\t\t->Fine doc $doc \n"
                    }
                    #print "\t->Fine Dest $dest \n";
                }
            }
        }
    }
}

print LISTA "+++ "  . getTimestamp() . " Fine elaborazione +++\n\n";

