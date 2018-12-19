#!/usr/local/bin/perl
use File::DirWalk;
use File::LineEdit;

        #my ($file, $ok);
        my $dw = new File::DirWalk;
	my $path = "";
	#	my $Ricevute	= "\\\\Webserver\\FileServer\\Entratel\\F24\\";
	my $Ricevute	= "\\\\Webserver\\FileServer\\Ricevute\\F24\\";
	#open(F24, ">\\\\Webserver\\FileServer\\F24\\F24.csv");
	open(F24, ">\\\\Webserver\\FileServer\\Ricevute\\F24\\F24.csv");
	#open(I24, ">\\\\Webserver\\FileServer\\F24\\I24.csv");
	open(I24, ">\\\\Webserver\\FileServer\\Ricevute\\F24\\I24.csv");

        $dw->onFile(sub {

                my ($file) = @_;
                my $ok   = 0;
                # elabora solo .rel
                if ( $file !~ /rel$/ ) {  return File::DirWalk::SUCCESS; }

		
                if ( $file =~ /F24\\2007/ ) {  $ok   = 1;}
                if ( $file =~ /F24\\2008/ ) {  $ok   = 1;}
                #print "++$ok++","\n" ;
                
                if ( $ok eq 0 ) {  return File::DirWalk::SUCCESS; }
		#print "++$file++","\n";

                my $rec = "";
                my $le = File::LineEdit->new($file, autosave=>0);

		my ($ProtF,$TipoF,$DataF,$FileF,$DataA);
		my ($NPagA,$NPagN,$NPagE,$NPagX);
                my ($CF, $Prot, $DataV, $Importo, $EsitoEl );
                my ($DataR, $NPagT, $NPagR, $NPagZ, $NPagS);
                my ($Banca, $EsitoPag,$Adj);
                
		$Prot= $CF= $DataV= $Importo= $EsitoPag= $Banca = $EsitoEl= "";
                $NPagT=$NPagR=$NPagZ=$NPagS= "";
		$ProtF=$TipoF=$DataF=$FileF= "";
		$DataA=$DataR=$NPagA=$NPagN=$NPagE=$NPagW=$NPagX= "";

                
		$file =~ s/\.rel$/\.pdf/;
		$file =~ s/^\\\\Webserver\\FileServer/http:\\\\Webserver/;
		$file =~ s/\\/\//g;
		$file = '"' . $file . '"';

		my $nRec = 0;
                foreach my $line (@$le)   {
			$nRec ++;
			#print $line, "\n";
                        if ( substr($line, 17, 1) eq "P" && substr($line, 27, 5) eq "F24A0" ) { # Record di testa
                              
			      $ProtF = '"' . substr($line,  0,17) . '"';	#Protocollo file
                              $TipoF = '"' . substr($line, 27, 5) . '"';	#Tipo file
                              $DataF = '"' . substr($line, 46, 8) . '"';	#Data Ricevimento
                              $FileF = '"' . substr($line, 68,24) . '"';	#File inviato

                              $NPagT = substr($line,751, 6);	#Pagamenti totali
                              $NPagR = substr($line,1160, 4);	#Ricevuti positivi
                              $NPagZ = substr($line,1473, 4);	#Ricevuti con versamento zero
                              $NPagS = substr($line,1622, 4);	#Scartati
                              
                         }
                        if ( substr($line, 17, 1) eq "R" && substr($line, 27, 5) eq "F24A0" ) { # Record di testa

			      #print $line, $rec, "\n";
			      
			      $Prot	=	substr($line,33, 5);		#Protocollo pagamento
			      $CF	= '"' . substr($line,38, 16) . '"';	#CF
			      if ($CF =~ /^"[0-9].*/ ) {
				      $CF =~ s/^"/"     / ; 
				      $CF =~ s/ *"$/"/ ; 
			      }
			      #$DataA	= 	substr($line,601,10) ;		#Data acquisizione
			      #$DataV	=	substr($line,988, 10) ;		#Data del versamento
			      #$Importo	=	substr($line,1071, 13);		#Importo versamento
			      #$EsitoEl	= '"' . substr($line,1228, 10) . '"';	#Esito elaborazione
			      #-------------
			      if ( $line =~ /Data versamento          : (..........)/) {   # Versamento nei termini
				      $DataV	= $1;
				      $DataA	= $1;
			      }
			      if ( $line =~ /Data versamento richiesta: (..........) .*Data versamento effettiva: (..........) /) {  
				      # Versamento fuori termine
				      $DataV	= $1;
				      $DataA	= $2;
			      }
			      if ( $line =~ /Importo versamento : E. *(............) .*Esito elaborazione       : (..........) /) {  
				# parse out values
				      $Importo	= $1;			
				      $EsitoEl	= '"' . $2 . '"';
				      $EsitoEl =~ s/\s+//g ;
				      $Importo =~ s/\s+//g ;
			      }
			      #------------
			      
			      $rec = join(';',$TipoF, $ProtF, $DataF, $FileF, $DataV, $DataA, $CF, $Prot, $Importo, $EsitoEl, 
				      		$NPagT, $NPagR, $NPagZ, $NPagS, $file, $nRec);
			      #$rec =~ s/ //g;
			      print F24 $rec, "\n";

                        }
                        
                         if (substr($line, 17, 1) eq "P"  && substr($line, 27, 5) eq "I24A0" ) { # Ricevuta file

			      #print $line, $rec, "\n";
			      
                              $ProtF = '"' . substr($line,  0,17) . '"';
                              $TipoF = '"' . substr($line, 27, 5) . '"';
                              $DataF = '"' . substr($line, 46, 8) . '"';
                              $FileF = '"' . substr($line, 68,24) . '"';

                              $DataA = substr($line,485,10);	#Data acquisizione
                              $NPagA = substr($line,681, 4);	#Pagamenti addebitati
                              $NPagN = substr($line,921, 4);	#Pagamenti non addebitati
                              $NPagE = substr($line,1161, 4);	#Pagamenti eseguiti
                              $NPagW = substr($line,1321, 4);	#Pagamenti in attesa
                              $NPagX = substr($line,1481, 4);	#Pagamenti annullati
                              
                        }
                        if ( substr($line, 17, 1) eq "R" && substr($line, 27, 5) eq "I24A0" ) { # Record di testa

				#print $line, $rec, "\n";
				
			      $Adj 	= 0;
			      $Prot	=	substr($line,33, 5);		#
			      $CF	= '"' . substr($line,38, 16) . '"';	#CF
			      if ($CF =~ /^"[0-9].*/ ) {
				      $CF =~ s/^"/"     / ; 
				      $CF =~ s/ *"$/"/ ; 
			      }
			      if ( substr($line,1441,4) eq "Data" ) { $Adj = 80;}

			      $Banca	= '"' . substr($line,1548+$Adj, 5) . '"';	#
			      $Importo	=	(substr($line,1468+$Adj, 18));	#
			      $EsitoPag = '"' . substr($line,1628+$Adj, 10) . '"';	#
			      if ($Importo =~ /  0,00$/) {
                                 $EsitoPag = '"' . substr($line,1548+$Adj, 10) . '"';#
                                 $Banca = "";
                              }
			      
			      $rec = join(';',$TipoF, $ProtF, $DataF, $FileF, $DataA, $CF, $Prot, $Importo, $EsitoPag, $Banca,
				     		 $NPagA, $NPagN, $NPagE, $NPagW, $NPagX, $file, $nRec);
			      #$rec =~ s/ //g;
                              print I24 $rec, "\n";
			      #print $rec, "\n";
                        }
                        if ( substr($line, 17, 1) eq "Q" && substr($line, 27, 5) eq "I24A0" ) { # Record di testa

				#print $line, $rec, "\n";

			      $Adj 	= 0;
			      if ( substr($line,1441,4) eq "Data" ) { $Adj = 80;}

			      $Prot	=	substr($line,33, 5);		#
			      $CF	= '"' . substr($line,38, 16) . '"';	#CF
			      if ($CF =~ /^"[0-9].*/ ) {
				      $CF =~ s/^"/"     / ; 
				      $CF =~ s/ *"$/"/ ; 
			      }
			      $Banca	= '"' . substr($line,1548+$Adj, 5) . '"';	#
			      $Importo	=	(substr($line,1468+$Adj, 18));	#
			      $EsitoPag = '"' . substr($line,1628+$Adj, 10) . '"';	#
			      if ($Importo =~ /0,00$/) {
                                 $EsitoPag = '"' . substr($line,1548+$Adj, 10) . '"';#
                                 $Banca = "";
                              }
			      
			      $rec = join(';',$TipoF, $ProtF, $DataF, $FileF, $DataA, $CF, $Prot, $Importo, $EsitoPag, $Banca,
				     		 $NPagA, $NPagN, $NPagE, $NPagW, $NPagX, $file, $nRec);
			      #$rec =~ s/ //g;
                              print I24 $rec, "\n";
			      #print $rec, "\n";
                        }

                }

		#close F24;
		#close I24;
                return File::DirWalk::SUCCESS;
        });
        # Home directory
        $dw->walk($Ricevute);


