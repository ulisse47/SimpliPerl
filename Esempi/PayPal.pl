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

my %opts = (
            help       => 0,
	    number     => 0,
	    update     => 0
            );

Getopt::Long::Configure('bundling');

GetOptions(
           'h|help'     => \$opts{help},
           'n|number'   => \$opts{number},
           'u|update'   => \$opts{update},
           'p|paese'    => \$opts{paese}
           ) ;
if ($opts{help}) {
   print "\nperl -w $0 [-h|-n|-u|-p]\n 
   	-u : Aggiorna le tabelle Pagamenti e Attivazioni
	-n : Elabora 1 solo messaggio 
	-p : Assegna paese di default ";
   exit 0;
}
if ($opts{update}) {
   print "\n+++ Aggiornamento DB attivato\n";
} else {
   print "\n+++ Aggiornamento DB disattivato\n";
}
my $paesedef = "";
if ($opts{paese}) {
   print "\n+++ Assegna paese di default ($paesedef) \n";
}
my $maxcount = 1;
if ($opts{number}) {
  	printf "+++ Opzione numero di messaggi da elaborare: $maxcount\n";	
  	print "+++ Elaborare 1 solo messaggio\n";	
}
#
#
# Prepara Outlook 
#
my $Outlook = Win32::OLE->new('Outlook.Application');
my $ol      = Win32::OLE::Const->Load($Outlook);
   Win32::OLE->Option( Warn => 1 );    #Ignora gli errori
my $namespace = $Outlook->GetNamespace("MAPI");
my $Folder    = $namespace->GetDefaultFolder(olFolderInbox);
my $Elaborati = $Folder->Folders->Item('Ricevute elaborate');
my $msg       = $Folder->Items;
my $count     = $msg->Count;
 
#
#
#	Variabili

my ( $ind,     $Paese,$err,     $cpmsg,$rc,   $risp );
my ( $sender,  $attach, $subject, $body, $date, $path );
my ( $CodTran, $DataOp, $Mese,    $Prod, $Addr, $User, $Ric );
my ( $Ndoc,    $Code,   $Email,   $Nome, $Cognome);

my $Tipo = 1; # Ricevuta pagamento PayPal
my $Pathlog = "T:/Log/";
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
#
#
# Elabora i messaggi in Posta in arrivo
#
#
open(LISTA,">>" . $Pathlog . "PayPal.log");
print LISTA "\n\nElaborazione Email PayPal pagamenti del $GG ore $HH\n *** $count Messaggi \n";

#printf ( LISTA "$opts{help}, $opts{number}, $opts{update}, $opts{paese} \n");
print "+++ Messaggi da elaborare $count - Data elab: $GG \n";

#
# Per ogni email ...
#
#

EMAIL: for ( $ind = $count ; $ind > 0 ; $ind-- ) {

    $DataOp = $CodTran = $Mese = $Prod = $Addr = $User = $Ric = $Code = $risp = $Paese = "";

    $subject = $msg->Item($ind)->Subject;
    $sender  = $msg->Item($ind)->SenderName;
    $body    = $msg->Item($ind)->Body;
    $date    = $msg->Item($ind)->SentOn;
    $attach  = $msg->Item($ind)->Attachments;

    printf( LISTA "\n1 +sender:	%s\n", $sender );

    if ( $sender =~ /(.+) tramite PayPal/ ) {    # parse
        $User = $1;
        ($Nome, $Cognome) = split(" ", $User);
    	#	scegli il paese
	$Paese = $paesedef;
	if (!$opts{paese}) {
		print "$Nome $Cognome e' italiano? [S|N]: ";
		$risp = getc(STDIN); 
		if ( $risp eq "S" ) {
			$Paese = 'Italy';
		}
	}
    } else {
	next;
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
            $Prod = $1;
            next;
        }
        if ( $line =~ /Ricevuta n�: ([\d-]+)/ ) {    # parse
            $Ric = $1;
            next;
        }
        if ( $line =~
             /hai ricevuto un pagamento di .50,00 USD da (.*)\. \((.*)\)/ )
        {                                            # parse
            $Addr = $2;
            next;
        }

    }
    #
    # Fine elaborazione messaggio
    #

    #
    # Lettura Numero di ricevuta disponibile
    #
    my $rf1 = $sth1->execute;
    if ($rf1) {
        my $rv1 = $sth1->bind_columns( \$Ndoc );
        if ( $rf1 = $sth1->fetch ) {
            if ( defined $Ndoc ) {
                printf( LISTA "3 +Numero ricevuta da attribuire: %s\n", $Ndoc + 1 );
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
    #
    # Ricerca Codice di attivazione
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
    #	Aggiornamento Tabella attivazioni
    #

    printf( LISTA
	"  ++Codice:	(%s)\n  ++DataOp:	(%s)\n  ++Prodotto:	(%s)\n" .
        "  ++User:	(%s)\n  ++Email:	(%s)\n  ++Ric. PP:	(%s)\n" .
	"  ++Nome: (%s)  ++Cognome: (%s)  ++Paese (%s)\n",
        $CodTran, $DataOp, $Prod, $User, $Addr, $Ric, $Nome, $Cognome, $Paese );

    #
    # Aggiornamento Attivazioni
    # Aggiornamento Pagamenti
    #
    if ($opts{update}) {

    	print "\n+++ Aggiornamento DB attivato\n";
    	$sth4->execute( $Addr, $Code )
      	or die "Errore in aggiornamento Attivazioni" . $DBI::errstr . "\n";

    	$sth2->execute( $DataOp, $Addr, $User, $Ndoc + 1, $CodTran, $Ric, $Prod,
                    $Code, $Nome, $Cognome, $Tipo, $Paese)
      	or die "!!! Errore su aggiornamento Pagamenti" . $DBI::errstr . "\n";
        printf( LISTA "5 +Database aggiornati\n" );

    #
    #	Sposta il messaggio nella cartella Elaborati
    #
    $cpmsg = $msg->Item($ind)->Move($Elaborati); 
    printf( LISTA "6 +Messaggio spostato su cartella Elaborati\n" );
    }
}

close(LISTA);

#  $rc= $dbh->commit     or die $dbh->errstr;
   $rc= $dbh->disconnect or warn $dbh->errstr;

