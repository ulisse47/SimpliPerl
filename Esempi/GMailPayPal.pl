#!/usr/bin/perl -w
#	Modifica modello di ricevuta PP
#       
#	Invia email ai destinatari estratti da un db

use strict;
use Getopt::Long;
use PDF::Reuse;
use File::Copy;
use DBI;
use Mail::Sender;

my %opts = (
            email      => 0,
            help       => 0,
            version    => 0,
            verbose    => 0,
	    numero     => 0
            );

Getopt::Long::Configure('bundling');
GetOptions('e|email'	=> \$opts{email},
           'h|help'     => \$opts{help},
           'V|version'  => \$opts{version},
           'n|numero'   => \$opts{numero},
           ) ;
if ($opts{help})
{
   print "perl -w $0 [-e|-V|-n] ";
   exit 0;
}

if ($opts{version})
{
   print "CAM::Reuse v$PDF::Reuse::VERSION\n";
   exit 0;
}

my @num;
if ($opts{numero})
{
   die "Ricevute da elaborare omesse" unless $#ARGV >= 0;
   my ($ind);
   for ( $ind=0; $ind<=$#ARGV; $ind++) {
   	 $num[$ind] = $ARGV[$ind];
	   }
   print "++ Parametri ricevute da elaborare comunicati: @num \n"; 
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

my $Tipo = "Pagamenti";      # Cartella degli incassi da PayPal

#
#  Impostazione ambiente
#

my $Path = "T:\\Ricevute\\";		# Home directory
#my $Path = "T:\\Tmp\\";

my $PathTipo = "T:\\Ricevute\\$Tipo\\";	# Archivio ricevute
stat $PathTipo || mkdir $PathTipo;

my $PathMod = "T:\\Modelli";			# Directory log
stat $PathMod   	|| die "$PathMod non esiste\n";

my $PathNew = $Path . "Ark\\$Tipo\\";			# Directory file elaborati
stat $Path . "Ark"  	|| mkdir $Path . "Ark";
stat $PathNew   	|| mkdir $PathNew;

my $PathLog = $Path . "Log\\$Tipo\\";			# Directory log
stat $Path . "Log"  	|| mkdir $Path . "Log";
stat $PathLog   	|| mkdir $PathLog;



chdir $Path || die "$Path non esiste\n" ;


my ($ndoc, $pag, $Periodo, $Anno, $Azienda, $Ditta, $p, $Pathfile);
my ($Pathdoc, $Cod, $Mitt, $Ufficio, $CF, $RagSoc);
my ($emailuff, %tabuff, %tabcnt, $ultric, $field );

my ($TipoDoc, $Data, $Tit, $Nome, $Cognome, $Email, $NR);
my ($Importo, $Prodotto, $Descrizione, $Transazione, $Ricevuta, $Div);
my ($Codice, $Paese);





open(LISTA,">" . $PathLog . "$Tipo.log");
print LISTA "+++ $tempo Inizio elaborazione +++\n*** Lista segnalazioni\n";

 
$Email = "mcavalieri47\@gmail.com";

	my $sender = Mail::Sender->new ({
    	from => "cavalierisrl\@ccprogetti.it",
	smtp => "authsmtp.ccprogetti.it",
	});
	$sender = $sender->new({on_errors => 'die'}); 
   	$sender = $sender->new({authid 	=> 'smtp@ccprogetti.it'}); 
   	$sender = $sender->new({authpwd   => 'Ccprogetti@2015'}); 

   	$rc = $sender->new({debug => 'T:/Bin/debug.txt'});

	if  (defined $Email and $opts{email}) {
		print "Prova ad inviare email\n";

		$sender->OpenMultipart( {	
			to      => "mcavalieri47\@ccprogetti.it",
			#to      => "$Email",
			subject => "Ricevuta di pagamento del prodotto n° 12345678",,
			port    => 25,
			priority=> 2,
			boundary => 'This-is-a-mail-boundary-435427'
  		});
		print "Apertura messaggio completata\n";

		$sender->Body({msg =>"Trasmettiamo il Codice di attivazione"});
		print "Corpo del messaggio completato\n";

		#$sender->Attach({
		#description 	=> 'Data Sheet',
		#ctype 		=> 'application/pdf',
		#encoding 	=> 'Base64',
		#file 		=> "T:/Ricevute/Pagamenti/RicPP_11.pdf",
		#disposition 	=> 'attachment; filename=*; type="PDF Document"'
		#});

		print "Messaggio inviato\n";
 		$sender->Close;	

} # Fine files da elaborare


printf LISTA "+++ $tempo Fine elaborazione +++\n\n";
