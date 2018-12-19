#!perl -w PayPal.pl
#PayPal.pl
#
#	Legge la posta nella cartella 'INPS' INPS.NotaRettifica@inps.it
#
#
#

use DBI;
use File::Copy;
use strict;
use Win32::OLE;
use Win32::OLE::Variant;
use Win32::OLE::Const 'Microsoft Outlook';
use Getopt::Long;
use PDF::Reuse;

my %opts = (
            help       => 0,
	    number     => 0,
	    update     => 0,
	    paese      => 0,
	    email      => 0
            );

Getopt::Long::Configure('bundling');

GetOptions(
           'h|help'     => \$opts{help},
           'n|number'   => \$opts{number},
           'u|update'   => \$opts{update},
           'p|paese'    => \$opts{paese},
           'e|email'    => \$opts{email}
           ) ;

my $paesedef = "Italy";
if ($opts{help}) {
   print "\nperl -w $0 [-h|-n|-u|-p-e]\n 
   	-h : Help
	-n : Elabora 1 solo messaggio 
	-p : Assegna paese di default ($paesedef)
   	-u : Aggiorna le tabelle Pagamenti e Attivazioni
	-e : Invia email
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
my $maxcount = 2;
if ($opts{number}) {
  	print "\n+++ Elaborare 1 solo messaggio\n\n";	
	$maxcount = 1;
}

#
# Prepara Outlook 
#
my $Outlook = Win32::OLE->new('Outlook.Application');
my $ol      = Win32::OLE::Const->Load($Outlook);
   Win32::OLE->Option( Warn => 2 );    #Ignora gli errori
my $namespace = $Outlook->GetNamespace("MAPI");
my $Folder    = $namespace->GetDefaultFolder(olFolderInbox);
my $Elaborati = $Folder->Folders->Item('Ricevute elaborate');
my $msg       = $Folder->Items;
my $count     = $msg->Count;


#
#	Definizione Variabili globali
#
my ( $ind,	$Paese,		$err,     $cpmsg,	$rc,   $risp );
my ( $sender,	$attach,	$subject, $body,	$date, $path );
my ( $CodTran,	$DataOp,	$Mese,    $Prodotto,	$Addr, $User, $Ricevuta );
my ( $Ndoc,	$Code,		$Email,   $Nome,	$Cognome);
my ( $modpdf,	$tiporic,	$Prezzo,  $Descrizione);
my ( $DocPdf,	$soggetto,	$corpo,   $modcorpo);

#
#  Impostazione ambiente
#

my $DirTipo = "Pagamenti";      # Tipo operazione: ricevuta di pagamento
my $Path = "T:/Ricevute/";		# Home directory

my $PathTipo = $Path . $DirTipo . "/";	# Archivio ricevute
stat $PathTipo || mkdir $PathTipo;

my $PathMod = "T:\\Modelli\\";			# Directory log
stat $PathMod   	|| die "$PathMod non esiste\n";

my $PathNew = $Path . "Ark\\$DirTipo\\";			# Directory file elaborati
stat $Path . "Ark"  	|| mkdir $Path . "Ark";
stat $PathNew   	|| mkdir $PathNew;

my $PathLog = $Path . "Log\\$DirTipo\\";			# Directory log
stat $Path . "Log"  	|| mkdir $Path . "Log";
stat $PathLog   	|| mkdir $PathLog;

################################################################################
#	Definione Percorsi e parametri
################################################################################
my $Tipo = 1;		# Ricevuta pagamento PayPal

my $Div	= "USD";	# Valuta utilizzata

my $Importo	= 50,00;	# Importo

if ( $count > $maxcount) {
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
# Prepara comandi SQL
#
my $DSN = "dbi:ODBC:CCP";
my $dbh = DBI->connect( $DSN, '', '' ) or die "$DBI::errstr\n";
my $sth1 = $dbh->prepare( " 
		SELECT 	max([Pagamenti].NR) 
		FROM [Pagamenti]
		GROUP BY [Pagamenti].Tipo; 
	") or die "$DBI::errstr\n";

my $sth2 = $dbh->prepare( " 
		INSERT INTO Pagamenti (Data, Email, RagioneSociale, NR,
		Transazione, Ricevuta, Prodotto, Codice, Nome, Cognome,
		Tipo, Paese) 
		VALUES (?,?,?,?,?,?,?,?,?,?,?,?);
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
open(LISTA,">>" . $PathLog . "PayPal.log");

print LISTA "\n\nElaborazione Email PayPal pagamenti del $GG ore $HH\n *** $count Messaggi \n";
print "+++ Messaggi da elaborare $count - Data elab: $GG \n";

    #
    # Lettura primo numero di ricevuta disponibile
    #
    my $rf1 = $sth1->execute;
    if ($rf1) {
        my $rv1 = $sth1->bind_columns( \$Ndoc );
        if ( $rf1 = $sth1->fetch ) {
            if ( defined $Ndoc ) {
                printf( LISTA "3 +Ultimo numero ricevuta utilizzato (%s)\n", $Ndoc );
		$rc  = $sth1->finish;
            }
        }
        else {
            die "$DBI::errstr\n - Errore su Pagamenti";
        }
    }
    else {
        die "$DBI::errstr\n - Errore su Pagamenti";
    }
####################################################################################
#
# 	Per ogni messaggio ricevuto ...
#
####################################################################################

EMAIL: for ( $ind = $count ; $ind > 0 ; $ind-- ) {

	$Ndoc += 1;
	printf( LISTA "3 ++Numero ricevuta in elaborazione (%s)\n", $Ndoc );

	$DataOp = $CodTran = $Mese = $Prodotto = $Addr = $User = $Ricevuta = $Code = $risp = $Paese = "";

	$subject = $msg->Item($ind)->Subject;
	$sender  = $msg->Item($ind)->SenderName;
	$body    = $msg->Item($ind)->Body;
	$date    = $msg->Item($ind)->SentOn;
	$attach  = $msg->Item($ind)->Attachments;

	printf( LISTA "\n1 +sender:	%s\n", $sender );

	if ( $sender =~ /(.+) tramite PayPal/ ) {    # parse
        	$User = $1;
        	($Nome, $Cognome) = split(" ", $User);
    	}

	printf( LISTA "2 +subject:	%s\n", $subject );
	if ( $subject =~ /Pagamento ricevuto da (.+)$/ ) {
		$Addr = $1;
	}

	#
	#Elabora il messaggio e rileva i dati
	#

	my @lines = split( '\n', $body );
	for my $line (@lines) {
        	chop $line;

        #printf( "+ + +>%s\n", $line);

        if ( $line =~
             /(..?) (...) (....) .* Codice transazione: (.................)/ )
        {    # parse
            $DataOp = join( "/", $1, $mesi{$2}, $3 );
            $CodTran = $4;
            next;
        }
        if ( $line =~ /Oggetto n. (\d+)/ ) {    # parse
            $Prodotto = $1;
            next;
        }
        if ( $line =~ /Ricevuta n°: ([\d-]+)/ ) {    # parse
            $Ricevuta = $1;
            next;
        }
        if ( $line =~
             /hai ricevuto un pagamento di .50,00 USD da (.*)\. \((.*)\)/ )
        {                                            # parse
            $Addr = $2;
            next;
        }

    } # Fine elaborazione messaggio

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
                    $Code, $Nome, $Cognome, $Tipo, $Paese)
      		or die "!!! Errore su aggiornamento Pagamenti" . $DBI::errstr . "\n";
        	printf( LISTA "6 +Database aggiornati\n" );

	}
	#
	#	Sposta il messaggio nella cartella Elaborati
	#
	$cpmsg = $msg->Item($ind)->Move($Elaborati); 
	printf( LISTA "5 ++Messaggio spostato su cartella Elaborati\n" );

	printf( LISTA
		"6 ++Codice:	(%s)\n6 ++DataOp:	(%s)\n6 ++Prodotto:	(%s)\n" .
        	"6 ++User:	(%s)\n6 ++Email:	(%s)\n6 ++Ric. PP:	(%s)\n" .
		"6 ++Nome:	(%s) ++Cognome: (%s)\n6 ++Paese:     (%s)\n",
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
        	printf( LISTA "7 ++Prodotto: (%s) \n\t\tPrezzo: (%s)\tValuta: (%s)\n",
			$Descrizione, $Prezzo, $Div );
	}
	$rc  = $sth5->finish;
	#
	#	Elaborazione del documento PDF
	# 
	if ($Paese eq 'Italy') {
		printf( LISTA "8 ++Modulistica italiana	\n" );
		$modpdf="pag_it.pdf";
		$tiporic="Ricevuta";
	} else {
		printf( LISTA "8 ++Modulistica inglese	\n" );
		$modpdf="pag_en.pdf";
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
       	prText(350,645, $Nome . " " . $Cognome);

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

	printf LISTA "9 ++ Documento PDF ($DocPdf) elaborato \n";



	#####################################################################
	#
	#	Invio messaggio con allegata ricevuta
	#
	#####################################################################
	

	#print "$Addr\n";
	#$Addr = 'maurizio.cavalieri@ccprogetti.it';
	print "$Addr\n";

	if  (defined $Addr and $opts{email}) {
		print "\n++++ Preparazione email in corso ... \n";
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
		#print $soggetto;
		open(BODY, "<$PathMod\\$modcorpo") or die "++ File body inesistente in $PathMod , $!";
		while(<BODY>){ $corpo = $corpo . $_; }
		close(BODY);		

		$corpo =~ s/---Nome---/$Nome/;
		$corpo =~ s/---Cognome---/$Cognome/;		
	

	my $newmsg = $Outlook->CreateItem(0) or die $!;
	$newmsg->{'To'}      = $Addr;
	$newmsg->{'Subject'} = $soggetto;
	$newmsg->{BodyFormat} = olFormatHTML;
	$newmsg->{'HTMLBody'} = "$corpo";
	$attach = $PathTipo . $DocPdf;
	#print "$attach\n";

	my $attachments = $newmsg->Attachments();
	$attachments->Add($attach);

	$newmsg->Send();

	my $error = Win32::OLE->LastError();
	printf(LISTA "10++Notifica inviata a $Addr \n") if not $error;
	printf(LISTA "10++Notifica non inviata \n") if $error;

	undef $corpo ;

}	# Fine messaggio
}

#
#  Fine programma
#

close(LISTA);
#$rc= $dbh->commit     or die $dbh->errstr;
$rc= $dbh->disconnect or warn $dbh->errstr;
