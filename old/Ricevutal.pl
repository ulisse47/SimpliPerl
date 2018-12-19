#!perl -w Ricevute.pl
#PayPal.pl
#
#	Legge le Notifiche PayPal di pagamento
#	Emette ricevuta
#	Aggiorna DB Access
#	Invia Email
#
#
#

use strict;
use DBI;
use File::Copy;
use Getopt::Long;
use PDF::Reuse;
use Mail::IMAPClient;
use Mail::Sender;
use IO::Socket;
use IO::Socket::SSL;
use DateTime::Format::Mail;
use Data::Dumper;
use Email::MIME::Encodings;
use HTML::Strip;	# Istanza HTML cleaner
	my $hs = HTML::Strip->new();
use Mail::IMAPClient::Bodystructure;
use Encode qw(encode decode);

my %opts = (
            help       => 0,
	    number     => 0,
	    update     => 0,
	    paese      => 0,
	    email      => 0,
	    move       => 0
            );

Getopt::Long::Configure('bundling');

GetOptions(
           'h|help'     => \$opts{help},
           'n|number'   => \$opts{number},
           'u|update'   => \$opts{update},
           'p|paese'    => \$opts{paese},
           'e|email'    => \$opts{email},
           'm|move'     => \$opts{move}
           ) ;

my $paesedef = "";
if ($opts{help}) {
   print "\nperl -w $0 [-h|-n|-u|-p|-e|-m]\n 
   	-h : Help
	-n : Elabora 1 solo messaggio 
	-p : Assegna paese di default ($paesedef)
   	-u : Aggiorna le tabelle Pagamenti e Attivazioni
	-e : Invia email
	-m : Sposta il messaggio
	";
   exit 0;
}

if ($opts{email}) {
   print "+++ Invio email attivato\n";
} else {
   print "+++ Invio email disattivato\n";
}

if ($opts{update}) {
   print "+++ Aggiornamento DB attivato\n";
} else {
   print "+++ Aggiornamento DB disattivato\n";
}
if ($opts{paese}) {
   print "+++ Assegna paese di default ($paesedef) \n";
}
if ($opts{move}) {
   print "+++ Archiviazione messaggio elaborato attivata \n";
}
if ($opts{email}) {
   print "+++ Invio messaggio con ricevuta attivato \n";
}

my $maxcount = 0;
if ($opts{number}) {
  	print "\n+++ Elaborare 1 solo messaggio\n\n";	
	$maxcount = 1;
}

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
	} else {
		$count = $client->message_count($folder);
		if( $count > 0) {
			print	"--- Trovati " . $count .  " messaggi nel folder: <$folder>\n" .
				"--- Account  <$mail_username>\n";
		} else { # No messages
			print "--- No messages to process. Bye\n";
			exit;
		}
	}
}

#
#	Definizione Variabili globali
#
my ( $ind,	$Paese,		$err,     $cpmsg,	$rc,   $risp );
my ( $CodTran,	$DataOp,	$Mese,    $Prodotto,	$User, $Ricevuta );
my ( $Ndoc,	$Code,		$Email,   $Nome,	$Cognome, $Addr);
my ( $modpdf,	$tiporic,	$Prezzo,  $Descrizione);
my ( $DocPdf,	$soggetto,	$corpo,   $modcorpo, 	$Pathdoc);

#
#  Impostazione ambiente per Ricevute PayPal
#
my $DirTipo = "Pagamenti";      # Tipo operazione: ricevuta di pagamento
my $Path = "T:/Ricevute/";	# Home directory

my $PathTipo = $Path . $DirTipo . "/";		# Archivio ricevute
stat $PathTipo || mkdir $PathTipo;

my $PathMod = "T:\\Modelli\\";			# Directory log
stat $PathMod   	|| die "$PathMod non esiste\n";

#my $PathNew = $Path . "Ark\\$DirTipo\\";	# Directory file elaborati
#stat $Path . "Ark"  	|| mkdir $Path . "Ark";
#stat $PathNew   	|| mkdir $PathNew;

my $PathLog = "T:/Log/";			# Directory log
stat $Path . "Log"  	|| mkdir $Path . "Log";
stat $PathLog   	|| mkdir $PathLog;

################################################################################
#	Definione Percorsi e parametri
################################################################################
my $Tipo	= 1;		# Ricevuta pagamento PayPal
my $Div		= "USD";	# Valuta utilizzata
my $Importo	= 50,00;	# Importo

if ( $count > $maxcount ) {
	$count = $maxcount;
}

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
		GROUP BY [Pagamenti].Tipo; 
	" ) or die "$DBI::errstr\n";

my $sth2 = $dbh->prepare( " 
		INSERT INTO Pagamenti (Data, Email, RagioneSociale, NR,
		Transazione, Ricevuta, Prodotto, Codice, Nome, Cognome,
		Tipo, Paese, Notifica) 
		VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?);
	") or die "$DBI::errstr\n";

my $sth3 = $dbh->prepare( " 
		SELECT 	[CodeFree].code
		FROM [CodeFree]
	") or die "$DBI::errstr\n";

my $sth4 = $dbh->prepare( " 
		UPDATE  Attivazioni 
		SET  [Attivazioni].status = '01',  [Attivazioni].address = ?
		WHERE  [Attivazioni].code = ?
	") or die "$DBI::errstr\n";

my $sth5 = $dbh->prepare( " 
		SELECT 	[Prodotti].Descrizione,	[Prodotti].Prezzo, [Prodotti].Valuta
		FROM [Prodotti]
		WHERE [Prodotti].Codice = ?
	") or die "$DBI::errstr\n";

#
# Apri file log
#
open(LISTA,">>" . $PathLog . $0 . ".log");

print LISTA "---------------------------------------------------------------\n" . 
		"Elaborazione Email PayPal pagamenti del $GG ore $HH\n\n";

#
# Lettura primo numero di ricevuta disponibile
#
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


####################################################################################
#
# 	Per ogni messaggio ricevuto ...
#
####################################################################################

my @search = ('SUBJECT "Notifica ricezione pagamento da" FROM "@paypal.it>"');
my @results = $client->search( @search );

@search = ('SUBJECT "Pagamento ricevuto da" FROM "@paypal.it>"');
push   @results , $client->search( @search );

#print join("\n",@results),"\n";
if (! @results ) { # Controllo contenuto di folder
	print "+-- $GG - Nessun messaggio da elaborare\n";
	exit 1;
} else {
	print "+-- $GG - Messaggi da elaborare ( " . ($#results + 1) . " )\n";
	my $n = 0;
	my ( $subject, $date, $sender, $body, $content , $uid) = undef;
	EMAIL: foreach(@results) {
		$uid = $_;

		#if ($uid != '1235') { next EMAIL;} # Debug

		my $struct = Mail::IMAPClient::BodyStructure->new($client->fetch($uid,"bodystructure")) ;
		my @h = split("\n" , Dumper($struct));
		foreach my $l (@h) {
			 if ($l =~ /\'bodyenc\' => \'(.+)\'/) { 
				$content = $1;
			}
		}

		$Ndoc += 1; $n += 1;
		if ($opts{number} and $maxcount < $n) { last EMAIL ;}

		printf( LISTA "\n------\n0 ++Numero ricevuta in elaborazione ( %s )\n", $Ndoc );
		printf( LISTA "0 ++Uid messaggio in elaborazione ( %s )\n", $uid );

		$DataOp = $CodTran = $Mese = $Prodotto = $Addr = $User = $Ricevuta = $Code = $risp = $Paese  = 
		$Nome = $Cognome = "";

		#
		# Estrazione dati da Header 
		#

		$subject = $client->subject( $uid )	or die "Could not Subject_string: ", $client->LastError;

		$date    = $client->date( $uid )	or die "Could not Date: ", $client->LastError; #print $date . "\n";
		$DataOp = DateTime::Format::Mail->parse_datetime( $date );
		$DataOp = $DataOp->dmy('/') . " " . $DataOp->hms ;
		#print $DataOp . "\n";


		$sender  = $client->get_header( $uid , "From" ) or die "Could not From: ", $client->LastError;
		$sender  = decode("MIME-Header", $sender);
		#print "\n " . $sender . "\n";
		printf( LISTA "\n1 +sender:	%s\n", $sender );
		if ( $sender =~ /[ "]*(.+) tramite PayPal/ ) {
        		$User = $1;
        		($Nome, $Cognome) = split(" ", $User);
			print "\n+-- User " .  $User . "\n";
    		}
		# Lettura email e codice prodotto
		$subject = decode("MIME-Header", $subject);
		printf( LISTA "2 +subject:	%s\n", $subject );

		if (  $subject =~ /Pagamento ricevuto da (.+)$/ ) {
			$Addr = $1;
			print "+-- Email " .  $Addr . "\n";
		}

		if ( $subject =~ /N° oggetto (\d+) - Notifica ricezione pagamento da .+ .+ \((.+)\)$/ ) {
			$Prodotto = $1;
			$Addr = $2;
			print "+-- Email " .  $Addr . "\n";
			print "+-- Prodotto " .  $Prodotto . "\n";
		}


		# Definizioni
		my $c_text = undef;
		my $dcd = undef;

		# Body
		if ($content eq "base64" ) {
			# Decode bodymsg
			$body = $client->bodypart_string( $uid,'1' )	or die "Could not get bodypart string: ", $client->LastError;
			$dcd = Email::MIME::Encodings::decode(base64 => $body);
			#print $content . "\n";
		} else {
			$dcd    = $client->body_string($uid)	or die "Could not body_string: ", $client->LastError;
		}

		#
		#Elabora il corpo del messaggio e rileva i dati
		#
		$c_text = $hs->parse( $dcd );
		$hs->eof;
		$c_text =~ tr/ / /s;

		#print  Dumper($c_text) ;
		#####################################################
		#	Elabora ogni linea del corpo del messaggio
		#####################################################
		my @lines = split( '\n', $c_text );
		my $i = undef;
		LINE:	for ($i=0; $i <= $#lines; $i++) {
			if ($lines[$i] =~ /^ *$/) { next LINE;}

			my $line = $lines[$i];
			#printf( "+++>%s\n", $line); # Debug
			# Lettura codice transazione 
			if ( $CodTran eq "" and $line =~ /Codice transazione: / ) {
				$line .= $lines[$i+1]; 
				if ($line =~ /Codice transazione: (.................)/) {
            				$CodTran = $1;
					print "+-- Codice transazione: " .  $CodTran . "\n";
            				next LINE;
				}
        		}	
			# Lettura N. ricevuta 
			if ( $Ricevuta eq "" and $line =~ /([\d-]+) Conserva / ) {
            			$Ricevuta = $1 ;
				print "+-- Codice Ricevuta: " .  $Ricevuta . "\n";
        		}	
        		if ($Ricevuta eq "" and  $line =~ /Ricevuta n°: ([\d-]+) / ) {  
            			$Ricevuta = $1;
        		}
			# Lettura codice prodotto 
			if ( $Prodotto eq "" and $line =~ /Oggetto n.. (\d+)/ ) {
            			$Prodotto = $1;
				print "+-- Prodotto " .  $Prodotto . "\n";
        		}
			# Lettura user 
			if ( $User eq '' and $line =~ /acquirente (.+) .*\@/ ) {
            			$User = $1;
				print "+-- User " .  $User . "\n";
				($Nome, $Cognome) = split(" ", $User);
        		}
    		} # Fine elaborazione messaggio di notifica

		#
		# Ricerca Codice di attivazione disponibile
		#
		$sth3->execute or die "Esecuzione lettura codice di attivazione fallita " .
    			$DBI::errstr . "\n";
		$sth3->bind_columns( \$Code ) or die $DBI::errstr;
		if ( $sth3->fetch ) {
        		printf( LISTA "4 +Codice attivazione:	(%s)\n", $Code );
			$rc  = $sth3->finish;
		}
		else {
        		printf( LISTA "\n!!!!! Codici di attivazione esauriti !!!!!\n" );
        		die "Codici di attivazione esauriti" . "\n";
		}
    
		#
		#	Aggiornamento Tabella 
		#

		#
		#	Scelta della lingua
		#

		if ($opts{paese}) {
			$Paese = $paesedef;
		} else {
			print "$Nome $Cognome e' italiano? [S|N]: ";
			while (defined($risp = getc (STDIN)) ) {
				if ($risp eq "s" or $risp eq "S"){
					print "Si\n";
					$Paese = 'Italy';
					last;
				}
				elsif	($risp eq "N" or $risp eq "n" ){
					print "No\n";
					$Paese = '';
					last;
				}
			}
		}



		#
		# Aggiornamento Tablle Attivazioni e Pagamenti
		#
		if ($opts{update}) {

    			print "\n+++ Aggiornamento DB attivato\n";
    			$sth4->execute( $Addr, $Code )
      				or die "Errore in aggiornamento Attivazioni" . $DBI::errstr . "\n";
	
    			$sth2->execute( $DataOp, $Addr, $User, $Ndoc, $CodTran, $Ricevuta, $Prodotto,
                    		$Code, $Nome, $Cognome, $Tipo, $Paese, $GG . " " . $HH)
      					or die "!!! Errore su aggiornamento Pagamenti" . $DBI::errstr . "\n";
        		printf( LISTA "6 +Database aggiornati\n" );

		}



	
		printf( LISTA
			"6 ++Transazione:	(%s)\n6 ++DataOp:	(%s)\n6 ++Prodotto:	(%s)\n" .
        		"6 ++User:	(%s)\n6 ++Email:	(%s)\n6 ++Ric. PP:	(%s)\n" .
			"6 ++Nome:	(%s) ++Cognome: (%s)\n6 ++Paese:      (%s)\n",
        		$CodTran, $DataOp, $Prodotto, $User, $Addr, $Ricevuta, $Nome, $Cognome, $Paese
		);


		#####################################################################
		#
		#	Preparazione Ricvuta PDF
		#
		#####################################################################
	
		#
		# Lettore tabella Prodotti
		#
		$sth5->execute($Prodotto) or die "Lettura Prodotto fallita " .
    			$DBI::errstr . "\n";
		$sth5->bind_columns( \$Descrizione,\$Prezzo,\$Div ) or die $DBI::errstr;
		if ( $sth5->fetch ) {
        		printf( LISTA "7 ++Prodotto:   (%s) \n\t\tPrezzo: (%s)\tValuta: (%s)\n",
			$Descrizione, $Prezzo, $Div );
		}
		$rc  = $sth5->finish;
		#
		#	Elaborazione del documento PDF
		# 
		if ($Paese eq 'Italy') {
			printf( LISTA "8 ++Modulistica italiana	\n" );
			$modpdf="PagCcp_it.pdf";
			$tiporic="Ricevuta";
		} else {
			printf( LISTA "8 ++Modulistica inglese	\n" );
			$modpdf="PagCcp_en.pdf";
			$tiporic="Payment Receipt";
		}

		prDocDir($PathTipo);
		$DocPdf ="RicPP_" . $year . "_" . $Ndoc . ".pdf" ;
	
		prFile($DocPdf);
		prFont('Times-Roman');   		# Just setting a font
		prCompress(1);     
		prForm($PathMod . "\\$modpdf");         # Here the template is used
		prFontSize(10);
    
		prText(340,660, "Spett.le");
       		prText(350,645, $User);
	
        	prText(350,630, $Addr);
        	prText(70,525,  $tiporic);
        	prText(270,525, $Ndoc . "\/" . $year);
        	prText(435,525, $DataOp);
        	prText(240,485, $Ricevuta);
        	prText(60,425,  $Prodotto);
        	prText(430,425, $Div);
        	prText(385,425, "1");
        	prText(510,425, $Prezzo);
	
		prFontSize(8);
        	prText(185,425, $Descrizione);
		prFontSize(12);
        	prText(210,180, $Code);
	
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
				$soggetto = "Ricevuta di pagamento del prodotto n. ";
				$modcorpo = "bodymsg_it.html";
			} else {
				$soggetto = "Payment Receipt - product n. ";
				$modcorpo = "bodymsg_en.html";
			}		
			$soggetto = $soggetto . " " . $Prodotto . " - " . $Descrizione;

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
					to	=>$Addr,
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
		#	Sposta il messaggio nella cartella Elaborati
		#
		if ($opts{move}) {
			$cpmsg = $client->move($arkfolder, $uid)or die "Could not move: $@\n";
			printf( LISTA "9 ++Messaggio spostato su cartella Elaborati\n" );
			printf( "+++Messaggio spostato su cartella $arkfolder\n" );
		}
	} # Fine messaggio
} # Fine Messaggi

#
#  Fine programma
#
#$rc= $dbh->commit     or die $dbh->errstr;
$client->expunge; 
$dbh->disconnect or warn $dbh->errstr;
printf LISTA "++  Fine elaborazione - $GG +++\n\n";
close(LISTA);
