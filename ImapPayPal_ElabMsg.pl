	EMAIL: foreach(@results) {

		#EMAIL: for ( $ind = $count ; $ind > 0 ; $ind-- ) {

		$Ndoc += 1;
		printf( LISTA "3 ++Numero ricevuta in elaborazione (%s)\n", $Ndoc );

		$DataOp = $CodTran = $Mese = $Prodotto = $Addr = $User = $Ricevuta = $Code = $risp = $Paese = "";

		$subject = $client->subject( $_ )or die "Could not Subject_string: ", $client->LastError;;
		$date    = $client->date( $_ ) or die "Could not Date: ", $client->LastError;
		$sender  = $client->get_header( $_, "From" ) or die "Could not From: ", $client->LastError;
		$body    = $client->body_string( $_ ) or die "Could not body_string: ", $client->LastError;
		#$attach  = $msg->Item($ind)->Attachments;

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
   }

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
	#$cpmsg = $msg->Item($ind)->Move($arkfolder); 
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
	
	}

	my $newmsg;
	#my $newmsg = $Outlook->CreateItem(0) or die $!;
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
