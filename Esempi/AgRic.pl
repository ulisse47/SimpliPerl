#!/usr/bin/perl -w
#	Elabora Ricevute Agenzia delle Entrate
#       Estrae ogni pagina ed la archivia in directory strutturate
#	Invia email ai destinatari estratti da un db

use strict;
use Getopt::Long;
use CAM::PDF;
use CAM::PDF::PageText;
use File::Copy;
use DBI;
use Mail::Sender;
use PDF::Reuse;

my %opts = (
            email      => 0,
            help       => 0,
            version    => 0,
            verbose    => 0,
            );

Getopt::Long::Configure('bundling');
GetOptions('e|email'	=> \$opts{email},
           'h|help'     => \$opts{help},
           'V|version'  => \$opts{version},
           'v|verbose'  => \$opts{verbose},
           ) ;
if ($opts{help})
{
   print "perl -w AgRic.pl [-e|-V] [MOD740 | MOD750] ";
   exit 0;
}

if ($opts{version})
{
   print "CAM::PDF v$CAM::PDF::VERSION\n";
   exit 0;
}

my ($sec, $min, $hour, $mday, $mon, $year, $wday, $yday, $isdst) = localtime(time);
$year += 1900;  $mon	+= 1;
my $tempo =   join('/', $mday,$mon,$year) . " " .  join(':', $hour,$min,$sec) ;

my %Month = (
		'GENNAIO','01',	'FEBBRAIO','02', 'MARZO','03',
		'APRILE','04',	'MAGGIO','05',	 'GIUGNO','06',
		'LUGLIO','07',	'AGOSTO','08',	 'SETTEMBRE','09',
		'OTTOBRE','10', 'NOVEMBRE','11', 'DICEMBRE','12'
		);

my $suf	 = "Vari";

# -------------------------- Impostazione ambiente -------------
#

#Server Smtp
my $smtp = "10.0.0.10";

#
# Home directory
#
#my $Path = "E:\\Fileserver\\Ricevute\\";
my $Path = "C:\\Tmp\\";
my $Tipo = "Vari";
my $PathTipo = $Path . $Tipo. "\\";
stat $PathTipo || mkdir $PathTipo;

# --------------------------------------------------------------

my $PathNew = $Path . "Ark\\$Tipo\\";			# Directory file elaborati
stat $Path . "Ark"  	|| mkdir $Path . "Ark";
stat $PathNew   	|| mkdir $PathNew;

my $PathLog = $Path . "Log\\$Tipo\\";			# Directory log
stat $Path . "Log"  	|| mkdir $Path . "Log";
stat $PathLog   	|| mkdir $PathLog;

my $DSN = "Caf";
my $dbh = DBI->connect("dbi:ODBC:$DSN", '', '');


chdir $Path || die "$Path non esiste\n" ;
my @Source = <$suf*.pdf>;
#print @Source, "\n";

my ($file, $Pathfile, $Codice);

my ($Mitt, $CF, $Pagina);
my ($uff, $emailuff, %tabuff, %tabcnt );

my $sender;
$sender = new Mail::Sender

#	Carica tabella uffici

my $sthu = $dbh->prepare(" SELECT Cod, Sede FROM Uffici Order by Cod; ");

my $rfu	= $sthu->execute();
	if ( $rfu) {
		my $rvu = $sthu->bind_columns(\$uff, \$emailuff );
		while ( $rfu = $sthu->fetch ) {
#                        printf ("\"%s\"; \"%s\"\n", $uff, $emailuff );
			$tabuff{$uff}=$emailuff;
#                        print $tabuff{$uff}, "\n";
			}
	}

my $sth = $dbh->prepare(
	" SELECT Contribuente, Gruppo, Codice FROM _730 WHERE (Codice = ? ) ; "
	);

open(LISTA,">" . $PathLog . "$Tipo.log");
print LISTA "+++ $tempo Inizio elaborazione +++\n*** Lista segnalazioni\n";

#
# Per ogni file di ricevute
#
my ( $p, $Pag );
for $file ( @Source ) {

(my $BFile, my $Suf) = split(/\./, $file);
#print $file, "\n";
my ($TotP, $str, $doc, $key, $depkey );
$depkey = "";

   $doc	= CAM::PDF->new($file); die "$CAM::PDF::errstr\n" if (!$doc);
   $TotP= $doc->numPages();		# Numero pagine nel file
   $TotP = 220;
   $doc->rangeToArray(1,$TotP,$p);

#   print $TotP, "\n";



#
#	************** Per ogni pagina **************
#
my ($Pathdoc, $Modello, $Anno, $Azienda, $Ufficio, $Mese, $DocName );
for ($p = 1 ; $p <= $TotP; $p =$p + 1) {
	#
	#	Elabora la pagina pdf ed estrae i codici
	#
   	$str= "";  $Ufficio = "0000"; $Pagina = "00"; $DocName ="";
        $Modello = "NN"; $Anno = "Anno"; $Mese = "Mese";
	$Azienda = ""; $CF = "CF"; $Codice = "----"  ;

        print $p,"-", $TotP, "-",$file, "\n";
	$str = $doc->getPageText($p);

        if (defined $str) {
	       $str =~ s/\x00//g;
        SWITCH: {
       	       if ($str =~ /SEGNALAZIONE ERRORI ELABORAZIONE MENSILE/) {
                  if ( $str =~ /(\d\d\d\d) P PAGHE/) { $Ufficio   = $1; }
                  if ( $str =~ /MOD.(...)        ([A-Z]*) +(\d\d) EL/) {
	                $Modello = $1; $Mese = $Month{$2}; $Anno = "20".$3;
                        }
                  $DocName = $Modello;
                  last SWITCH;
                  }
               if ($str =~ /DA VERSARE ENTRO IL \d\d.(\d\d).(\d\d\d\d)/) {
                  $Modello = "F24"; $Mese = $1; $Anno = $2;
                  if ( $str =~ /\s?.. (\d\d\d\d) ([A-Z]) (\d\d\d\d)\n/) {
                        	$Ufficio = $1; $Azienda =  $1.$2.$3;
                   	}
                  $DocName = $Modello;
                  #print LISTA $Modello,$Mese,$Anno,$Azienda, "\n", $str, "\n";
	          last SWITCH;
                  }
               if ($str =~ / SCADENZA IL \d\d.(\d\d).(\d\d\d\d) .* TELEMATICA /) {
	          $Modello = "F24T"; $Mese = $1; $Anno = $2;
                  if ( $str =~ /\s?.. (\d\d\d\d) ([A-Z]) (\d\d\d\d)\n/) {
                        	$Ufficio = $1; $Azienda =  $1.$2.$3;
                   	}
                  $DocName = $Modello;
                  #print LISTA $Modello, $Mese, $Anno, $Azienda, "\n", $str, "\n";
                  last SWITCH;
	          }
               if ($str =~ /Riepilogo per ufficio dei modelli elaborati/) {
	          $Modello = "nnn";
                  print LISTA $Modello, $Mese, $Anno, $Azienda, "\n", $str, "\n";
                  last SWITCH;
	          }
               if ($str =~ /(\d\d\d\d) P ([A-Z])(\d\d\d\d) .* SITUAZIONE INDENNITA\' FINE RAPPORTO/) {
                  $Ufficio = $1; $Azienda =  $1.$2.$3;
                  if ( $str =~ /MOD.(...)        ([A-Z]*) +(\d\d) EL/) {
	                $Modello = $1; $Mese = $Month{$2}; $Anno = "20".$3;
                        }
                  $DocName = $Modello;
                  #print LISTA $Modello, $Mese, $Anno, $Azienda, "\n", $str, "\n";
                  last SWITCH;
	                }
               if ($str =~ /(\d\d\d\d) P ([A-Z])(\d\d\d\d) .* RILEVAZIONE DEL COSTO DEL PERSONALE/) {
	          $Ufficio = $1; $Azienda =  $1.$2.$3;
                  if ( $str =~ /MOD.(...)        ([A-Z]*) +(\d\d) EL/) {
	                $Modello = $1; $Mese = $Month{$2}; $Anno = "20".$3;
                        }
                  $DocName = $Modello;
                  #print LISTA $Modello, $Mese, $Anno, $Azienda, "\n", $str, "\n";
                  last SWITCH;
	                }
               if ($str =~ /(\d\d\d\d) P ([A-Z])(\d\d\d\d) .* ANALISI VERSAMENTI UNIFICATI/) {
                  $Ufficio = $1; $Azienda =  $1.$2.$3;
                  if ( $str =~ /MOD.(...)        ([A-Z]*) +(\d\d) EL/) {
	                $Modello = $1; $Mese = $Month{$2}; $Anno = "20".$3;
                        }
                  $DocName = $Modello;
                  #print LISTA $Modello, $Mese, $Anno, $Azienda, "\n", $str, "\n";
                  last SWITCH;
	          }
               if ($str =~ /PERIODO PAGA DI/ ) {
                  $Modello = "Ced";
                  if ( $str =~ /    (\d\d\d\d) ([A-Z])(\d\d\d\d)\n/) {
                        	$Ufficio = $1; $Azienda =  $1.$2.$3;
                   	}
                  if ( $str =~/                  ([A-Z]+) +(\d\d)\n/) {
                                $Mese = $Month{$1}; $Anno = "20".$2;
                        }
                  $DocName = $Modello;
                  #print LISTA $Modello, $Mese, $Anno, $Azienda, "\n", $str, "\n";
                  last SWITCH;
                  }
               if ($str =~ /Data autoriz. 08.01.200[89]/ ) {
                  $Modello = "LUn";
                  if ( $str =~/\s*([A-Z]{6}\d\d[A-Z]\d\d[A-Z]\d\d\d[A-Z])/) {
                        	$CF = $1;
                     }
                  if ( $str =~ /    (\d\d\d\d) ([A-Z]\d\d\d\d)\n/) {
                        	$Ufficio = $1; $Azienda = $1.$2;
                   	}
                  if ( $str =~/\s([A-Z]*)\s+(\d\d)\n/) {
                                $Mese = $Month{$1}; $Anno = "20".$2;
                     }
                  $DocName = $Modello;
                  #print LISTA $Modello, $Mese, $Anno, $Azienda,$CF, "\n", $str, "\n";
                  last SWITCH;
                  }
               if ($str =~ /(\d\d\d\d) ([A-Z]) (\d\d\d\d).*SCARICO PER IMPRESA CEDOLINI UTILIZZATI/) {
                  $Modello = "VID";
                  $Ufficio = $1; $Azienda =  $1.$2.$3;
                  if ( $str =~ /.*  MOD.VID   NEL MESE (\d\d).(\d\d)/) {
                        $Mese = $1; $Anno = "20".$2;
                  }
                  $DocName = $Modello;
                  #print LISTA $Modello, $Mese, $Anno, $Azienda, $CF, "\n", $str, "\n";
                  last SWITCH;
	          }
               if ($str =~ /DATI PER PROSP\. INF\. PERSONALE /) {
                  $Modello = "P57";
                  if ($str =~ /(\d\d\d\d) P ([A-Z])(\d\d\d\d) /) {
                        $Ufficio = $1; $Azienda =  $1.$2.$3;
                        }
                  if ( $str =~/MOD.P57\s+([A-Z]+)\s+(\d\d)\s/) {
                                $Mese = $Month{$1}; $Anno = "20".$2;
                  }
                  $DocName = $Modello;
                  #print LISTA $Modello, $Mese, $Anno, $Azienda,$CF, "\n", $str, "\n";
                  last SWITCH;
	                }
               if ($str =~ / *(\d\d\d\d) ([A-Z]) (\d\d\d\d) \d\d \d\d \d\d\d \d \d\n/) {
	          $Modello = "Mal";
                  $Ufficio = $1; $Azienda =  $1.$2.$3;
                  if ( $str =~/                  ([A-Z]+)\s+(\d\d\d\d)\n/) {
                                $Mese = $Month{$1}; $Anno = $2;
                        }
                  $DocName = $Modello;
                  #print LISTA $Modello, $Mese, $Anno, $Azienda,$CF, "\n", $str, "\n";
                  last SWITCH;
	          }
               if ($str =~ /     (\d\d\d\d) ([A-Z]) (\d\d\d\d) \d\d \d\d \d\d\d \d\s+(\d\d).(\d\d\d\d)\s+/) {
                  $Modello = "Liq";
                  $Ufficio = $1; $Azienda =  $1.$2.$3; $Mese = $4; $Anno = $5;
                  if ( $str =~/                  ([A-Z]+)\s+(\d\d\d\d)\n/) {
                                $Mese = $Month{$1}; $Anno = $2;
                        }
                  $DocName = $Modello;
                  #print LISTA $Modello, $Mese, $Anno, $Azienda,$CF, "\n", $str, "\n";
                  last SWITCH;
                  }
               $Modello = "NN";
               print LISTA  $str, "\n";
       	  }
	}
        $key = $Anno.$Mese.$Modello.$Azienda;

        #print LISTA $key, "-", $depkey, "\n";
        if ( $key eq $depkey ) {
             $Pag += 1;

        }
        else {
             $Pag = 1 ;
             $depkey = $key;
        }
        #$DocName = $DocName . "_" . $Pag;
        printf LISTA ("%s;%s;%s;%s;%s;%s;%s;%s \n", $Modello, $Mese, $Anno, $Azienda, $Pag, $key, $depkey, $p);
        $depkey = $key;

        my ( $Ditta, $Uff );
	my $rf	= $sth->execute($Codice);
	if ( $rf) {
		my $rv = $sth->bind_columns(\$Ditta, \$Uff, \$CF );
		if ( $rf = $sth->fetch ) {
			printf LISTA ("%s %-30s %s\n",$Uff, $Ditta,  $CF );

			$tabcnt{$Ufficio} +=1;
			#print 	"$Ufficio, $tabcnt{$Ufficio}\n";
		}
		else {
			#printf LISTA "*** Azienda inesistente: $Codice\n";
#                        $Ufficio = "XXXX";
		}

#		$Pathdoc  = $Pathfile . "\\" . "$Codice-$Azienda" . ".pdf";
#		unlink $PathTipo . $Anno . "\\" . $Ufficio . "\\" . "\\_tutti.pdf";
	}
	else {
                print LISTA "***Errore database - $Codice \n";
        }
#
# 	Archivio documento
#
	$Pathfile = $PathTipo . $Anno . "\\" . $Ufficio . "\\" . $Mese ;
	if ($Azienda ne "") {
              	$Pathfile = $Pathfile  . "\\" . $Azienda;
                }
	if (-d $Pathfile) { # La directory esiste
        }
        else {
              mkdir $PathTipo . $Anno  ;
              mkdir $PathTipo . $Anno . "\\" .  $Ufficio;
              mkdir $PathTipo . $Anno . "\\" .  $Ufficio . "\\" . $Mese;
              mkdir $Pathfile  unless -d $Pathfile;

        }

	prDocDir("$Pathfile");
	prFile("\\$DocName.pdf");
	if ( $Pag == 1 ) {
        	prDoc($file, $p, $p);
        } else {
                prDoc($file, $p-$Pag+1, $p);
        }
	prEnd();


} # Fine pagine file

copy($file,$PathNew . $file) or die "Copy failed: $!";
#unlink($file) or die "Unlink failed: $!";
print "$file elaborato \n";

} # Fine files da elaborare

print LISTA "\n--- Riepilogo per ufficio dei modelli elaborati\n";

#while (my ($key, $val) = each %tabcnt) { # Invio messaggio all'ufficio
foreach my $key (sort(keys %tabcnt)) {
	#print $key, '=', $tabcnt{$key}, "\n";
	#print "$key = $val \n"; print "$tabuff{$key} \n";
	my $val = $tabcnt{$key};
        printf LISTA ("%s:%5s $Tipo\n", $key , $val);
#        if  (defined $tabuff{$key} and $opts{email}) {
#        $sender->Open( {
#                        to      => "Maurizio.Cavalieri\@unioneartigiani.it",
                        #to     => $tabuff{$key},
#                        from    => "Info.Ced\@unioneartigiani.it",
#                        subject => "Ricevute INPS",
#                        smtp    => "$smtp",
#                        port    => 25,
#                        priority=> 2,
#                });
#        $sender->SendLineEnc("Ufficio $key: Sono stati elaborate n. $val ricevute $Tipo");
#        $sender->Close;
#
#        }
#               print LISTA "$tabuff{$key} - spedito \n";
}
print LISTA "+++ $tempo Fine elaborazione +++\n\n";
