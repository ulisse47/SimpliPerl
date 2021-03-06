#!perl -w PayPal.pl
#PayPal.pl
#
#	Legge ultimo record DB
#	Emette ricevuta
#	Invia Email
#
#

use strict;
use DBI;
use File::Basename;
use Getopt::Long qw(GetOptions);
use PDF::Reuse;
use Mail::IMAPClient;
use Mail::Sender;
use IO::Socket;
use IO::Socket::SSL;
#use DateTime::Format::Mail;
#use Data::Dumper;
use Email::MIME::Encodings;
use HTML::Strip;	# Istanza HTML cleaner
	my $hs = HTML::Strip->new();
use Mail::IMAPClient::Bodystructure;
use Encode qw(encode decode);

my %opts = (
            help       => 0,
	    email      => 0,
	    num	       => 0
            );

Getopt::Long::Configure('bundling');

GetOptions(
           'h|help'     => \$opts{help},
           'e|email'    => \$opts{email},
           'n|num=n'	=> \$opts{num}
           ) ;

my $paesedef = "";
if ($opts{help}) {
   print "\nperl -w $0 [-h|-e|-n]\n 
   	-h : Help
	-e : Invia email
	-n : Numero ricevuta da elaborare
	";
   exit 0;
}

if ($opts{email}) {
   print "+++ Invio email attivato\n";
} else {
   print "+++ Invio email disattivato\n";
}

if ($opts{num}) {
	print "+++ Elaboro la ricevuta n. $opts{num} \n";
} else {
	print "+++ Elaboro l'ultima ricevuta\n"
}

#exit;

#
# Config  imap connction
#
my $mail_hostname	= 'pop.ccprogetti.it';
my $mail_hostname_s	= 'pop.securemail.pro';
my $mail_username	= 'amministrazione@ccprogetti.it';
my $mail_password	= 'ccprogetti';
my $mail_ssl = 1;

#
# Set folders
#
my $folder	= 'INBOX';
my $arkfolder	= 'INBOX.PayPal';

#
# Make sure this is accessable for this namespace
#
my $socket = undef;

if( $mail_ssl ) {
	# Open up a SSL socket to use with IMAPClient later
	$socket = IO::Socket::SSL->new(
		PeerAddr => $mail_hostname_s,
		PeerPort => 993,
		Timeout => 5,
	);
} else {
	# Open up a none SSL socket to use with IMAPClient later
	$socket = IO::Socket::INET->new(
		PeerAddr => $mail_hostname,
		PeerPort => 143,
		Timeout => 5,
	);
}

# Check we connected
if( ! defined( $socket ) ) {
	print STDERR "Could not open socket to mailserver: $@\n";
	exit 1;
}

my $client = Mail::IMAPClient->new(
	Socket   => $socket,
	User     => $mail_username,
	Password => $mail_password,
	Timeout => 5,
);

# Check we have an imap client
if( ! defined( $client ) ) {
	print STDERR "Could not initialize the imap client: $@\n";
	exit 1;
}

# Check we are authenticated
my $count = undef;
if( $client->IsAuthenticated() ) {
	# Select the INBOX folder
	if( ! $client->select($folder )) {
		print STDERR "Could not select the INBOX: $@\n";
		exit;
	}
}

#
#	Definizione Variabili globali
#
my ( $ind,	$Paese,		$err,     $cpmsg,	$rc,   $risp );
my ( $CodTran,	$DataOp,	$Mese,    $Prodotto,	$User, $Ricevuta );
my ( $Ndoc,	$Code,		$Email,   $Nome,	$Cognome, $Addr);
my ( $modpdf,	$tiporic,	$Prezzo,  $Descrizione,	$Div, $Tit, $DescProd	);
my ( $DocPdf,	$soggetto,	$corpo,   $modcorpo, 	$Pathdoc, $Nr, $Date);

#
#  Impostazione ambiente per Ricevute PayPal
#
my $DirTipo = "Pagamenti";      # Tipo operazione: ricevuta di pagamento
my $Path = "T:/Ricevute/";	# Home directory

my $PathTipo = $Path . $DirTipo . "/";		# Archivio ricevute
stat $PathTipo || mkdir $PathTipo;

my $PathMod = "T:\\Modelli\\";			# Directory log
stat $PathMod   	|| die "$PathMod non esiste\n";

my $PathLog = "T:/Log/";			# Directory log
#stat $Path . "Log"  	|| mkdir $Path . "Log";
stat $PathLog   	|| mkdir $PathLog;
$PathLog   = $PathLog . basename($0);
stat $PathLog   	|| mkdir $PathLog;

################################################################################
#	Definione Percorsi e parametri
################################################################################

#
# Calcolo data
#
my ( $sec, $min, $hour, $mday, $mon, $year, $wday, $yday, $isdst ) =
  localtime(time);
$mon  = $mon + 1;
$mday = $mday;
$year = $year + 1900;
my $data = $year . substr( "00" . $mon, -2 ) . substr( "00" . $mday, -2 );

my $GG = join( '/', $mday, $mon, $year );
my $HH = join( ':', $hour, $min, $sec );

my %mesi = (
             "gen", "01", "feb", "02", "mar", "03", "apr", "04",
             "mag", "05", "giu", "06", "lug", "07", "ago", "08",
             "set", "09", "ott", "10", "nov", "11", "dic", "12"
);

#
#	Prepara email
#
my $send = Mail::Sender->new ({
    		from => 'amministrazione@ccprogetti.it',
		smtp => 'authsmtp.ccprogetti.it',
		port => 25,
});
$send = $send->new({on_errors	=> 'die'}); 
$send = $send->new({authid	=> 'smtp@ccprogetti.it'}); 
$send = $send->new({authpwd	=> 'Ccprogetti@2015'}); 
$send = $send->new({debug	=> 'T:/Log/debug.log'});
$send = $send->new({auth	=> 'PLAIN'});

#
# Prepara comandi SQL
#
my $DSN = "dbi:ODBC:CCP";
my $dbh = DBI->connect( $DSN, '', '' ) or die "$DBI::errstr\n";
my $sth1 = $dbh->prepare( " 
		SELECT 	max([Pagamenti].NR) 
		FROM [Pagamenti]
		WHERE [Pagamenti].Tipo = 0 
		; 
	" ) or die "$DBI::errstr\n";

my $sth2 = $dbh->prepare( " 
		SELECT 
		TipoDoc.Descrizione, Pagamenti.Data, Pagamenti.Tit, 
		Pagamenti.Nome, Pagamenti.Cognome, Pagamenti.Email, 
		Pagamenti.NR, Prodotti.Prezzo AS Importo, Pagamenti.Prodotto, Prodotti.Descrizione, 
		Pagamenti.Ricevuta, Prodotti.Valuta AS Div, 
		Pagamenti.Codice, Pagamenti.Paese 
		FROM ( Pagamenti INNER JOIN Prodotti ON Pagamenti.Prodotto = Prodotti.Codice) 
			INNER JOIN TipoDoc ON Pagamenti.Tipo = TipoDoc.Codice		
		WHERE NR = ?
	") or die "$DBI::errstr\n";

#
# Apri file log
#
print  $PathLog ."\n";
open(LISTA,">>" . $PathLog . ".log") or die "Open failed \n";
print LISTA "---------------------------------------------------------------\n" . 
		"Elaborazione Email PayPal pagamenti del $GG ore $HH\n\n";

#
# Selta del numero di ricevuta disponibile
#
if ($opts{num}) {
	$Ndoc =  $opts{num} ;
	print "+++ Elaboro la ricevuta n. $Ndoc \n";
} else {
	print "+++ Elaboro l'ultima ricevuta\n";
	my $rf1 = $sth1->execute;
	if ($rf1) {
        	my $rv1 = $sth1->bind_columns( \$Ndoc );
        	if ( $rf1 = $sth1->fetch ) {
            	if ( defined $Ndoc ) {
                	printf( LISTA "0 +Ultimo numero ricevuta utilizzato (%s)\n", $Ndoc );
			$rc  = $sth1->finish;
            	}
        	} else {
            	die "$DBI::errstr\n - Errore su Pagamenti";
		}
	} else {
        	die "$DBI::errstr\n - Errore su Pagamenti";
	}

}

	#####################################################################
	#
	#	Preparazione Ricvuta PDF
	#
	#####################################################################

	#
	# Lettore tabella query RicevutePP_Q
	#
	$sth2->execute($Ndoc) or die "Lettura RicevutePP fallita " .
    			$DBI::errstr . "\n";
	$sth2->bind_columns( 
			\$Descrizione, \$Date, \$Tit, \$Nome, \$Cognome, \$Addr,
			\$Nr, \$Prezzo, \$Prodotto, \$DescProd, \$Ricevuta, \$Div,
			\$Code, \$Paese
	) or die $DBI::errstr;
	if ( $sth2->fetch ) {
        		printf( LISTA "7 ++Prodotto:   (%s) \n\t\tPrezzo: (%s)\tValuta: (%s)\n",
			$DescProd, $Prezzo, $Div );
	}
		$rc  = $sth2->finish;
		#
		#	Elaborazione del documento PDF
		# 
		if ($Paese eq 'Italy') {
			printf( LISTA "8 ++Modulistica italiana	\n" );
			$modpdf="PagCcpEval_it.pdf";
			$tiporic="Evaluation";
		} else {
			printf( LISTA "8 ++Modulistica inglese	\n" );
			$modpdf="PagCcpEval_en.pdf";
			$tiporic="Evaluation";
		}

		prDocDir($PathTipo);
		$DocPdf ="__EvlPP_" . $year . "_" . $Ndoc . ".pdf" ;
		$Prezzo ="0,00";
	
		prFile($DocPdf);
		prFont('Times-Roman');   		# Just setting a font
		prCompress(1);     
		prForm($PathMod . "\\$modpdf");         # Here the template is used
		prFontSize(10);
    
		prText(340,660, "Spett.le");
		$User = $Tit . " " . $Nome . " " . $Cognome;
       		prText(350,645, $User);
	
        	prText(350,620, $Addr);
        	prText(70,525,  $tiporic);
        	prText(270,525, $Ndoc . "\/" . $year);

		#print $Date . "\n";
		($DataOp, $sec) = split( ' ', $Date);
		print $DataOp . "\n";
		($year, $mon, $mday) = split('-', $DataOp);
		$DataOp = $mday . "/" . $mon . "/" . $year;
		print $DataOp . "\n";



        	prText(435,525, $DataOp);
        	prText(240,485, $Ricevuta);
        	prText(60,425,  $Prodotto);
        	prText(430,425, $Div);
        	prText(385,425, "1");
        	prText(510,425, $Prezzo );
	
		prFontSize(8);
        	prText(185,425, $DescProd);
		prFontSize(12);
        	prText(210,170, $Code);
		#prText(210,160, "SMY183GIWO8XHB7FHEDX");

	
        	prPage();
		prEnd();

		printf LISTA "9 ++Documento PDF ($DocPdf) elaborato \n";



		#####################################################################
		#
		#	Invio messaggio con allegata ricevuta
		#
		#####################################################################
		#$Addr = 'maurizio.cavalieri@ccprogetti.it';
		print "$Addr\n";

		if  (defined $Addr and $opts{email}) {
			printf "\n+++ Preparazione email in corso ... \n";
			#
			#	Carica modello html del messaggio nella lingua richiesta
			#
			if ($Paese eq 'Italy') {
				$soggetto = "Ricevuta di pagamento del prodotto  ";
				$modcorpo = "bodymsg_it.html";
			} else {
				$soggetto = "Payment Receipt - product n. ";
				$modcorpo = "bodymsg_en.html";
			}		
			$soggetto = $soggetto . " " . $Prodotto . " - " . $DescProd;

			#print $soggetto . "\n;
			open(BODY, "<$PathMod\\$modcorpo") 
				or die "+++ File body inesistente in $PathMod , $!";
			while(<BODY>){ $corpo = $corpo . $_; }
			close(BODY);		

			$corpo =~ s/---Nome---/$Nome/;
			$corpo =~ s/---Cognome---/$Cognome/;

			#
			# Prepara il messaggio da inviare
			#
			$send->OpenMultipart({
					to	=>"$Addr, $mail_username",
					subject	=>$soggetto
				});
			$send->Body( {
        				charset =>	'utf8', #iso-8859-15
        				encoding =>	'7BIT',
        				ctype =>	'text/html',
					msg =>		"$corpo"
			});
			printf LISTA "10++ Corpo del messaggio per $Nome $Cognome completato\n";

			my $allegato = $PathTipo . $DocPdf;
			$send->Attach({
				description 	=> 'Data Sheet',
				ctype 		=> 'application/pdf',
				encoding 	=> 'Base64',
				file 		=> $allegato,
				disposition 	=> 'attachment; filename=*; type="PDF Document"'
			});

			printf LISTA	"11++ Messaggio inviato\n";
			printf 		"+++ Messaggio inviato\n";
 			$send->Close;	
	
			undef $corpo ;
	
   		} #Fine invio email
				

#
#  Fine programma
#
#$rc= $dbh->commit     or die $dbh->errstr;
$client->expunge; 
$dbh->disconnect or warn $dbh->errstr;
printf LISTA "++  Fine elaborazione - $GG +++\n\n";
close(LISTA);
