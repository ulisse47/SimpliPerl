#!perl -w Inps.pl
#Inps.pl
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

     my $DSN = "LibriPaga";
     my $dbh = DBI->connect("dbi:ODBC:$DSN", '', '');

     my $Outlook	= Win32::OLE->new('Outlook.Application');
     my $ol		= Win32::OLE::Const->Load($Outlook);
        Win32::OLE->Option( Warn => 1); 			#Ignora gli errori
     my $namespace	= $Outlook->GetNamespace("MAPI");
     my $Folder		= $namespace->GetDefaultFolder(olFolderInbox);
     my $Archivio 	= $Folder->Folders->Item('INPS');		#Cartella di servizio
     my $Spediti 	= $Archivio->Folders->Item('SPEDITI');		#Cartella di servizio
     my $msg		= $Archivio->Items;
     my $count		= $msg->Count;
 
     my ($ind, $work, $err, $cpmsg);
     my ($sender, $attach, $subject, $body, $date, $path, $pathsped ) ;
     my ($Matricola, $Anno, $Mese, $Prog);
     my ($Impresa, $Ufficio, $Email);
	
     # Calcolo data
     my ($sec,$min,$hour,$mday,$mon,$year,$wday,$yday,$isdst) = localtime(time);
	$mon	=$mon + 1;
	$mday	=$mday;
	$year	=$year + 1900;
	my $data= $year . substr("00" . $mon,-2 ) .  substr("00" . $mday,-2 );

	my $GG	= join('/',$mday,$mon,$year);
	my $HH	= join(':',$hour,$min,$sec) ;

	#$path		= "\\\\Webserver\\fileserver\\INPS\\Msg\\" ;
	$path		= "" ;

     #
     # Elabora tutti i messaggi in Posta in arrivo 
     #

     #print "*** $count \n";
     open(LISTA,">>" . $path . "Inps.log");
     print LISTA "Elaborazione NotaRettifica del $GG ore $HH\n *** $count \n";

 EMAIL: for ($ind = $count; $ind > 0; $ind--) {

                $subject	= $msg->Item($ind)->Subject;
                $sender		= $msg->Item($ind)->SenderName;
                $body		= $msg->Item($ind)->Body;
                $date		= $msg->Item($ind)->SentOn;
		$attach		= $msg->Item($ind)->Attachments;

		if (	$subject =~ /^Ricevuta/i || $subject =~ /^Non recapitato/i) {
			# Elabora solo i messaggi validi
			next;
			}

		if ( $sender !~ /NotaRettifica/) {	# Elabora solo i messaggi INPS.NotaRettifica@inps.it
			next;
			}

			#printf( LISTA "---- %s +++\n", $sender ); 

		#
		#	Elabora il messaggio e copia le denunce 
		#	

		my @lines = split('\n',$body);
		for my $line ( @lines ) {chop $line;

			#printf( "%s\n", $line);
			$Matricola = $Anno = $Mese = $Prog = "";

			if ( $line !~ /per il periodo/) {next};

			if ( $line =~ /---> (..........) per il periodo (..).(....)    Codice Sede: (....)/) {   # parse 
       			$Matricola	= $1;
        		$Anno		= $2;
        		$Mese		= $3;
        		$Prog		= $4;
    			}else {
				printf( LISTA "\nErrore : %s  \n", $subject);
			}


			#printf( LISTA "Matricola:%s Anno:%s  Mese:%s Sede:%s\n", $Matricola, $Anno, $Mese, $Prog);


			my $sth = $dbh->prepare(" 
				SELECT 	[MatricoleInps].Matricola, [MatricoleInps].Impresa, 
					[MatricoleInps].Ufficio, [MatricoleInps].EMail
				FROM [MatricoleInps]
				WHERE ([MatricoleInps].Matricola LIKE \'$Matricola%\' ) ; " 
			);  
			my $rf	= $sth->execute;   
			if ( $rf) { 

				my $rv = $sth->bind_columns(\$Matricola, \$Impresa, \$Ufficio, \$Email );

				if ( $rf = $sth->fetch ) {	# Matricola in archivio Unione 
					if (defined $Email ) { 
						printf( LISTA "%s %s %s %s %s %s %s\n", 
						$Matricola, $Anno, $Mese, $Impresa, $Ufficio,$Email, $date );
					}
		   			$body = "-- COMUNICAZIONE INVIATA DALL\'INPS PER LA DITTA $Impresa --" . $body;
				} 
			else {	# Matricola inesistente
        			print LISTA "$Matricola: ***Matricola Inesistente \n";
		   		$body = "-- MATRICOLA NON PRESENTE IN ARCHIVIO --\n\n" . $body;
				$Email = "Nessuno";
			}
			} 
			else {
        			print LISTA "$Matricola: ***Errore database \n";
				$err ++;
			}
		}
		   $body		= "-- Messaggio elaborato automaticamente da Unione Servizi --\n\n" . $body;
		my $ItNew		= $Outlook->CreateItem(olMailItem);
		   $ItNew->{Subject}	= $subject;
		   $ItNew->{Body}	= $body;
		   #$Email		= "ufficio.telematico\@unioneartigiani.it";
	 	my $work		= $ItNew->Recipients->Add($Email);
		   if ( $Email  ne "Nessuno" ){	
			   $work  = $ItNew->Send;
			   $cpmsg = $msg->Item($ind)->Move($Spediti);#Sposta il messaggio nella cartella di servizio
		   }
		#printf( LISTA "\n\n+++ %s\n\n\n",$subject);
 
	}

     close(LISTA);
