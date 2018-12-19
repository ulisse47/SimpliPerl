#!/usr/bin/perl -w
#	Elabora file di Bollettini Freccia
#       Estrae ogni pagina ed la archivia in directory strutturate
#	Invia email ai destinatari estratti da un db

use strict;
use Getopt::Long;
use CAM::PDF;
use File::Copy;
use DBI;
use Mail::Sender;

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
   print "perl -w $0 [-e|-V] ";
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
		'gennaio','01',	'febbraio','02', 'marzo','03',
		'aprile','04',	'maggio','05',	 'giugno','06',
		'luglio','07',	'agosto','08',	 'settembre','09',
		'ottobre','10', 'novembre','11', 'dicembre','12'
		);

my $Tipo = "CEMI";      # Deve essere CEMI
die "Tipo non ammesso; CEMI" unless $Tipo eq "CEMI" ;
#print $Tipo, "\n";

#
#  Impostazione ambiente
#
my $smtp = "smtp.fastwebnet.it";

my $Path = "E:\\Fileserver\\Ricevute\\";		# Home directory
#my $Path = "C:\\Tmp\\";
my $PathTipo = "E:\\Fileserver\\Ricevute\\$Tipo\\";	# Archivio ricevute
#my $PathTipo = "C:\\Tmp\\$Tipo\\";
stat $PathTipo || mkdir $PathTipo;

my $PathNew = $Path . "Ark\\$Tipo\\";			# Directory file elaborati
stat $Path . "Ark"  	|| mkdir $Path . "Ark";
stat $PathNew   	|| mkdir $PathNew;

my $PathLog = $Path . "Log\\$Tipo\\";			# Directory log
stat $Path . "Log"  	|| mkdir $Path . "Log";
stat $PathLog   	|| mkdir $PathLog;

my $DSN = "Intranet";                                   # DSN ODBC
my $dbh = DBI->connect("dbi:ODBC:$DSN", '', '');


chdir $Path || die "$Path non esiste\n" ;
my @Source = <$Tipo*.pdf>;
#print @Source;

my ($file, $Periodo, $Anno, $Azienda, $Ditta, $p, $Pathfile);
my ($Pathdoc, $Email, $Cod, $Mitt, $Ufficio, $CF, $RagSoc);
my ($uff, $emailuff, %tabuff, %tabcnt );

my $sender;
$sender = new Mail::Sender;

#	Carica tabella uffici

my $sthu = $dbh->prepare("
		SELECT 	Codice, Email FROM Uffici ; "
		);
my $rfu	= $sthu->execute();
	if ( $rfu) {
		my $rvu = $sthu->bind_columns(\$uff, \$emailuff );
		while ( $rfu = $sthu->fetch ) {
			#printf ("\"%s\"; \"%s\"\n", $uff, $emailuff );
			$tabuff{$uff}=$emailuff;
			#print $tabuff{$uff}, "\n";
			}
	}

#	Prepara query per ricerca ditta

my $sth = $dbh->prepare("
		SELECT 	CodPag, Emailx, Email, Cod, CF, RagSoc
                FROM emaildossier
                WHERE (RagSoc like ? ) ; "
		);

open(LISTA,">>" . $PathLog . "$Tipo.log");
print LISTA "+++ $tempo Inizio elaborazione +++\n*** Lista segnalazioni\n";

#
#	Elabora file in home directory
#

for $file ( @Source ) {		# Per ogni file

(my $BFile, my $Suf) = split(/\./, $file);
#print $file, "\n";


my $doc	= CAM::PDF->new($file); die "$CAM::PDF::errstr\n" if (!$doc);
my $TotP= $doc->numPages();	# Numero pagine nel file
#print $TotP, "\n";

#
#	Per ogni pagina
#

for ($p = 1; $p <= $TotP; $p++) {
	#
	#	Elabora la pagina pdf ed estrae il titolare
	#

	$doc = CAM::PDF->new($file);
	#print $p,"-", $TotP, "-",$file, "\n";
	$doc->rangeToArray(1,$TotP,$p);
	my $str = $doc->getPageText($p);

	if (defined $str)
	{
         CAM::PDF->asciify(\$str);
	 $str =~ s/\n//g;
	#	 print LISTA $str, "\n";


	if ( $str =~ /PERIODO (\d\d)\/(\d\d\d\d) .*Debitore (.*) Cinimporto/ ) {   # parse
		  $Anno		= $2;
		  $Periodo	= $1;
                  $Azienda	= $3;
		}
	# 	print LISTA "$Azienda-$Periodo-$Anno-\n";

	}

	#
	#	Cerca l'email dell'azienda
	#
	$Cod =	substr($Azienda,0,20);
        $Cod =~	s/\s*$//   ;
	my $rf	= $sth->execute($Cod . "%");
	if ( $rf) {
		my $rv = $sth->bind_columns(\$Ditta, \$Email, \$Mitt, \$Ufficio, \$CF, \$RagSoc );
		if ( $rf = $sth->fetch ) {
                       printf LISTA (
   			"\"%s\"; \"%s\"; \"%s\"; \"%s\"; \"%s\"\n",
                       	$Ditta, $Email, $Mitt, $Ufficio, $CF
			);

			if (defined $Email and $Email  ne "" ) {	# Azienda con email
				$Pathfile = $PathTipo . $Anno . "\\" . $Ufficio . "\\" . $Periodo;
				$Pathfile = $Pathfile . "\\Email\\" ;
			} else {				#Azienda senza email
				$Pathfile = $PathTipo . $Anno . "\\" . $Ufficio . "\\" . $Periodo;
				$Pathfile = $Pathfile . "\\Stampa\\" ;
				undef ($Email);
			}


			#
			#	Salva la pagina della ditta
			#
#	 		$doc->extractPages($p);
			$Pathdoc = $Pathfile . "\\" . $Ditta . ".pdf";
#			$doc->cleanoutput($Pathdoc);
			unlink $Pathfile . "\\_tutti.pdf";
			$tabcnt{$Ufficio} +=1;
			#print 	"$Ufficio, $tabcnt{$Ufficio}\n";
		}
		else {
			print LISTA "*** Azienda inesistente: $Cod\n";
			undef ($Email);
                        $Ufficio = "0000";
                        $Pathfile = $PathTipo . $Anno . "\\" . "$Ufficio" . "\\" . $Periodo;
                        $Pathdoc  = $Pathfile . "\\" . "$Azienda" . ".pdf";
		}
	}
	else { print LISTA "*** Errore database - $Ditta\n"; }
#
# 	Archivio documento
#
	if (-d $Pathfile) { # La directory esiste
        }
        else {
              mkdir $PathTipo . $Anno  ;
              mkdir $PathTipo . $Anno . "\\" .  $Ufficio ;
              mkdir $PathTipo . $Anno . "\\" .  $Ufficio . "\\" . $Periodo ;
              if ($Ufficio ne "0000") { # La directory esiste
                  mkdir $PathTipo . $Anno . "\\" .  $Ufficio . "\\" . $Periodo . "\\Email"  ;
                  mkdir $PathTipo . $Anno . "\\" .  $Ufficio . "\\" . $Periodo . "\\Stampa"  ;
              }
        }
        $doc->extractPages($p);
        $doc->cleanoutput($Pathdoc);
#
# 	Invia email all'utente
#

	if  (defined $Email and $opts{email}) {
	$sender->OpenMultipart( {	
			#to      => "unione.servizi\@unioneartigiani.it",
			to      => "$Email",
			from    => "$Mitt",
			subject => "$Ditta - $Tipo $Periodo $Anno",,
			smtp    => "$smtp",
			port    => 25,
			priority=> 2,
			boundary => 'This-is-a-mail-boundary-435427'
  		});
	$sender->Body({msg =>"Trasmettiamo il Bollettino Freccia inviato da $Tipo relativo al periodo $Periodo $Anno\n\nUnione Servizi S.r.l."});
	$sender->Attach({
		description 	=> 'Data Sheet',
		ctype 		=> 'application/pdf',
		encoding 	=> 'Base64',
		file 		=> "$Pathdoc",
		disposition 	=> 'attachment; filename=*; type="PDF Document"'
		});
 	$sender->Close;
	
	#print LISTA "$Email - spedito \n";
   	}
} # Fine pagine file
copy($file,$PathNew . $file) or die "Copy failed: $!";
unlink($file) or die "Unlink failed: $!";
print "$file elaborato \n";

} # Fine files da elaborare

print LISTA "\n--- Lista uffici elaborati\n";
while (my ($key, $val) = each %tabcnt) { # Invio messaggio all'ufficio

	#print "$key = $val \n"; print "$tabuff{$key} \n";

	printf LISTA ("%s: Elaborati %5s bollettini\n", $key,$val);
	if  (defined $tabuff{$key} and $opts{email}) {
		$sender->Open( {	
			to   	=> $tabuff{$key},
			#to      => "unione.servizi\@unioneartigiani.it",
			from    => "Info.Ced\@unioneartigiani.it",
			subject => "Bollettini Freccia - Cassa edile Milano",
			smtp    => "$smtp",
			port    => 25,
			priority=> 2,
  		});
	$sender->SendLineEnc("Ufficio $key: Sono stati elaborati n. $val Bollettini $Tipo");
 	$sender->Close;
   	}
}
 print LISTA "+++ $tempo Fine elaborazione +++\n\n";
