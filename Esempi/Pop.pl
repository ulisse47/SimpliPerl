	#!perl -w pop.pl
	#pop.pl
	#
	#	Legge la posta pop nella cartella 
	#	
	#	
	#
	# Always be safe
	#use strict;
	use warnings;
	
	use Mail::POP3Client;
 	use Email::MIME::Encodings;


	my ($fn, $decoded);

	my $pop = new Mail::POP3Client( USER     => 'cavalierisrl@ccprogetti.it',
	                               	PASSWORD => 'Ulisse@1947',
	                               	HOST     => "pop.ccprogetti.it:110" );
	  
	#print 	$pop->POPStat(), "\n";
	#print 	$pop->Count(), "\n";
	#print 	$pop->Last(), "\n";

	for( my $i = 1; $i <= $pop->Count(); $i++ ) {
	    foreach( $pop->Head( $i ) ) {
		    #
		    #/^(From|Subject):\s+/i && print $_, "\n";
		    #
		    #/^Subject: Pagamento ricevuto da/i &&  $pop->BodyToFile( $i, $fh ) ;
		    if (/^Subject: Pagamento ricevuto da/i) {

			   	print $i, "\n" ;

				$fn = $pop->Body( $i );
  				$decoded = Email::MIME::Encodings::decode(base64 => $fn);
				print $decoded, "\n";

		  	}
	    }
	  }
	  $pop->Close();
	
