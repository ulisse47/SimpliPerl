#!/usr/bin/perl -w
#	Modifica modello di ricevuta PP
#       
#	Invia email ai destinatari estratti da un db

use strict;
use Getopt::Long;
use PDF::Reuse;
use File::Copy;
use DBI;
use Email::Sender;
#
# Definizione variabili globali
#

#
# Impostazione Dta e Ora
#
my ($sec, $min, $hour, $mday, $mon, $year, $wday, $yday, $isdst) = localtime(time);
   $year += 1900;  $mon	+= 1;
my $tempo =   join('/', $mday,$mon,$year) . " " .  join(':', $hour,$min,$sec) ;

my %Month = (
		'gennaio','01',	'febbraio','02', 'marzo','03',
		'aprile','04',	'maggio','05',	 'giugno','06',
		'luglio','07',	'agosto','08',	 'settembre','09',
		'ottobre','10', 'novembre','11', 'dicembre','12'
		);
#
#  Impostazione ambiente
#

my $Tipo = "Pagamenti";      # Tipo operazione: ricevuta di pagamento
my $Path = "T:/Ricevute/";		# Home directory

my $PathTipo = $Path . $Tipo . "/";	# Archivio ricevute
stat $PathTipo || mkdir $PathTipo;

my $PathMod = "T:\\Modelli\\";			# Directory log
stat $PathMod   	|| die "$PathMod non esiste\n";

my $PathNew = $Path . "Ark\\$Tipo\\";			# Directory file elaborati
stat $Path . "Ark"  	|| mkdir $Path . "Ark";
stat $PathNew   	|| mkdir $PathNew;

my $PathLog = $Path . "Log\\$Tipo\\";			# Directory log
stat $Path . "Log"  	|| mkdir $Path . "Log";
stat $PathLog   	|| mkdir $PathLog;

#
# Connessione A DB
#
my $DSN = "dbi:ODBC:CCP";                                   # DSN ODBC
my $dbh = DBI->connect("$DSN", '', '');
$dbh->{RaiseError} = 1; # do this, or check every call for errors

chdir $Path || die "$Path non esiste\n" ;
#
#
#	Variabili globali
#

my ($ndoc, $pag, $soggetto, $Anno, $corpo, $Ditta, $p, $Pathfile);
my ($Pathdoc, $Cod, $Mitt, $Ufficio, $CF, $RagSoc, $allegato);
my ($emailuff, %tabuff, %tabcnt, $ultric, $field, $modpdf );

my ($TipoDoc, $Data,	 $Tit, $Nome, $Cognome, $Email, $NR);
my ($Importo, $Prodotto, $Descrizione, $Transazione, $Ricevuta, $Div);
my ($Codice,  $Paese,	 $body);
my ( @num );
#
#	Log 
#
open(LISTA,">" . $PathLog . "$0.log");
print LISTA "\nElaborazione $0 del $tempo \n";

#
# Controllo parametri
#
my %opts = (
            email      => 0,
            help       => 0,
            version    => 0,
	    ricevute   => 0
            );

Getopt::Long::Configure('bundling');
GetOptions(
	   'e|email'	=> \$opts{email},
           'h|help'     => \$opts{help},
           'V|version'  => \$opts{version},
           'r|ricevute' => \$opts{ricevute},
           'u|ultima'   => \$opts{ultima},
           't|tutte'    => \$opts{tutte},
           ) ;

if ($opts{help})
{
   print "\nperl -w $0 [-e|-V|-u|-t|-r] [r1] [r2] ...\n
  	-e : Abilita invio email di notifica\n
	-r : Numeri ricevuta da elaborare.\n
	-V : Versione di PDF::Reuse
	-u : Elabora ultima ricevuta
	-t : Elabora tutte le ricevute senza notifica
	" ;
   exit 0;
}

if ($opts{email})
{
   print "++ Invio email attivato\n";
} else {
   print "++ Invio email disattivato\n";
}

if ($opts{version})
{
   print "++ CAM::Reuse v$PDF::Reuse::VERSION\n";
   exit 0;
}

if ($opts{ricevute})
{
	@num = ();
	die "++ Ricevute da elaborare omesse. Elaborazione terminata " unless $#ARGV >= 0;
	my ($ind);
	for ( $ind=0; $ind<=$#ARGV; $ind++) {
   	 	$num[$ind] = $ARGV[$ind];
	}
	print "++ Ricevute da elaborare @num \n"; 
}

if ($opts{tutte}) 
{
	@num = ();
	my $stht = $dbh->prepare("
		SELECT [Pagamenti].NR 
		FROM [Pagamenti]
		WHERE [Pagamenti].Notifica  is NULL; 
	") or die "$DBI::errstr\n";

	my $rct  = $stht->execute()  or die $DBI::errstr;
	   $rct  = $stht->bind_columns(\$NR) or die $DBI::errstr;
	while( $stht->fetch ) {
		push @num, $NR;
		#print "$NR \n";
	}
	$rct  = $stht->finish;
	print ("++ Lista ricevute: (@num)\n");
	printf( LISTA "++ Ricevute presenti su DB da notificare: (@num)\n" );
}
if ($opts{ultima})
{
	@num = ();
	print ("++ Selezionata elaborazione solo ultima ricevuta\n");
	my $sthu = $dbh->prepare("
		SELECT max([Pagamenti].NR) 
		FROM [Pagamenti]
		GROUP BY [Pagamenti].Tipo; 
	") or die "$DBI::errstr\n";

	my $rcu	= $sthu->execute()  or die $DBI::errstr;
	   $rcu = $sthu->bind_columns(\$NR);
	   $rcu = $sthu->fetch;
	   $rcu = $sthu->finish;
       	   printf( LISTA "++ Ricevuta da notificare: (%s)\n", $NR);
	   push @num, $NR;
	   print ("++ Ultima ricevuta: (@num)\n");
}
#
#	Prepara email
#
my $sender = Email::Abstract->new ({
    		from => 'cavalierisrl@ccprogetti.it',
		smtp => 'authsmtp.ccprogetti.it',
		port => 25,
});
$sender = $sender->new({on_errors => 'die'}); 
$sender = $sender->new({authid    => 'smtp@ccprogetti.it'}); 
$sender = $sender->new({authpwd   => 'Ccprogetti@2015'}); 
$sender = $sender->new({debug	  => 'T:/Bin/debug.txt'});
$sender = $sender->new({auth	  => 'PLAIN'});

#	Prepara query per ricerca cliente

my $sth = $dbh->prepare("
	SELECT TipoDoc.Descrizione, Pagamenti.Data, Pagamenti.Tit, 
		Pagamenti.Nome, Pagamenti.Cognome, Pagamenti.Email, Pagamenti.NR, 
		Pagamenti.Importo, Pagamenti.Prodotto, [Prodotti PP].Descrizione, 
		Pagamenti.Transazione, Pagamenti.Ricevuta, Pagamenti.Div, 
		Pagamenti.Codice, Pagamenti.Paese
	FROM ([Pagamenti] 
	INNER JOIN [Prodotti PP] ON Pagamenti.Prodotto = [Prodotti PP].Codice) 
	INNER JOIN TipoDoc ON Pagamenti.Tipo = TipoDoc.Codice
	WHERE Pagamenti.NR = (?)
;");
#
# Prepara update data notifica
#
my $sth1 = $dbh->prepare( " 
		UPDATE  Pagamenti 
		SET  [Pagamenti].Notifica = ?
		WHERE  [Pagamenti].NR = ?
	") or die "$DBI::errstr\n";

#
#	Ciclo elaborazione ricevute
#	-----------------------------
#	Per ogni ricevuta da elaborare 
#	1. Carica Modello
#	2. Legge tabella Pagamenti
#	3. Riempie il modello
#	4. Spedisce notifica
#	5. Salva
#
#
my @Source = @num;

for $ndoc ( @Source ) {		# Per ogni cwricevuta

	printf(LISTA "+++ Ricevuta in elaborazione n.: %s\n", $ndoc);
	my $rc = $sth->execute($ndoc)  or die $DBI::errstr;
	   $rc = $sth->bind_columns(
		\$TipoDoc, \$Data, \$Tit, \$Nome, \$Cognome, \$Email, \$NR, 
		\$Importo, \$Prodotto, \$Descrizione, \$Transazione, \$Ricevuta, \$Div, 
		\$Codice, \$Paese
	   );
	$rc = $sth->fetch  or die $DBI::errstr;
	$rc = $sth->finish;
	printf(LISTA "+++ Letto numero: %s - Codice: %s - Paese: (%s)\n",
		$NR, $Codice,  ($Paese // '')); 
	#
	# Aggiusta il formato data
	#
	my @dataf  = split (/ /, $Data, 7);
	my @dataf2 = split (/-/, $dataf[0], 7);
	my $Dataf3 = join('/',$dataf2[2] , $dataf2[1], $dataf2[0]);
	#print $dataf3;
	#
	#
	#	Scelta della lingua
	#
	if( ($Paese // '') eq "Italy" ) {
		$modpdf="pag_it.pdf";
	} else {
		$modpdf="pag_en.pdf";
	}
	#
	#	Elaborazione del documento PDF
	# 
	prDocDir($PathTipo);
	$Pathdoc ="RicPP_" . $year . "_" . $ndoc . ".pdf" ;
	prFile($Pathdoc);
	prFont('Times-Roman');   		# Just setting a font
	prCompress(1);     
	prForm($PathMod . "\\$modpdf");         # Here the template is used
	prFontSize(10);
    
	prText(340,660, "Spett.le");

	if ( defined $Tit ) {
        	prText(350,645, $Tit . " " . $Nome . " " . $Cognome);
	} else {
        	prText(350,645, $Nome . " " . $Cognome);
	}

        prText(350,630, $Email);
        prText(70,525,  $TipoDoc);
        prText(240,525, $NR . "\/" . $year);
        prText(410,525, $Dataf3);
        prText(240,485, $Ricevuta);
        prText(50,425,  $Prodotto);
        prText(425,425, $Div);
        prText(500,425, $Importo);

	prFontSize(8);
        prText(185,425, $Descrizione);
	prFontSize(12);
        prText(210,180, $Codice);

        prPage();
	prEnd();

	printf LISTA "+++ Documento PDF ($Pathdoc) elaborato \n";

	#
	# Invia messaggio con allegato
	#
	# 
	if  (defined $Email and $opts{email}) {
		print "++++ Preparazione email in corso ... \n";
		#
		#	Carica modello html del messaggio nella lingua richiesta
		#
		if(($Paese // '') eq "Italy" ) {
			$corpo = "bodymsg_it.html";
			$soggetto = "Ricevuta di pagamento del prodotto n. ";
		} else {
			$corpo = "bodymsg_en.html";
			$soggetto = "Payment Receipt - product n. ";
		}	

		$sender->OpenMultipart( {	
			#to      => "$Email",
			to      => 'cavalierisrl@ccprogetti.it', 
			#bcc     => 'assistenza@ccprogetti.it',
			subject => "$soggetto $Prodotto - $Descrizione",
			priority=> 2,
			boundary => 'This-is-a-mail-boundary-435427'
  		});
		print "++ Apertura messaggio completata\n";


		open(BODY, "<$PathMod\\$corpo") or die "++ File body inesistente in $PathMod , $!";
		while(<BODY>){ $body = $body . $_; }
		close(BODY);

		$body =~ s/---Nome---/$Nome/;
		$body =~ s/---Cognome---/$Cognome/;		

		$sender->Body( {
				msg =>		"$body",
        			charset =>	'iso-8859-15',
        			encoding =>	'7BIT',
        			ctype =>	'text/html',
				});
		print "++ Corpo del messaggio per $Nome $Cognome completato\n";

		my $allegato = $PathTipo . $Pathdoc;
		$sender->Attach({
			description 	=> 'Data Sheet',
			ctype 		=> 'application/pdf',
			encoding 	=> 'Base64',
			file 		=> $PathTipo . $Pathdoc,
			disposition 	=> 'attachment; filename=*; type="PDF Document"'
		});

		print "++ Messaggio inviato\n";
 		$sender->Close;	
		undef $body ;

		#
		# Aggiorna data invio notifica
		#
    		print "\n++ Aggiornamento DB con data di notifica\n";
    		$sth1->execute( $tempo, $NR ) or die
			"++ Errore in aggiornamento Pagamenti" . $DBI::errstr . "\n";
        	printf( LISTA "++ Database aggiornato\n" );

	} # Fine files da elaborare
}

printf LISTA "++  Fine elaborazione - $tempo +++\n\n";
