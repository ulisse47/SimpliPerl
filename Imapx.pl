#!perl -w Imax.pl
#Imax.pl
#
#	Legge la posta IMAP nella cartella 
#	
#	
#
# Always be safe
use strict;
use warnings;

# Use the module

use Net::IMAP::Client;;
use Data::Dumper;
use Email::MIME::Encodings;

	my $imap = Net::IMAP::Client->new(
	   	server	 => 'pop.ccprogetti.it',
       		username => 'eshop@ccprogetti.it',
      		password => 'ccprogetti',	   
		port	 => 143
		) or die "Could not connect to IMAP server";
	

	$imap->login( 'eshop@ccprogetti.it','ccprogetti') or
		die('Login failed: ');

	my $folder = 'INBOX';

	my $nm = $imap->examine($folder) or 
		die "Could not select: " . $@ . "\n";

        # get list of folders
	my @folders = $imap->folders;
	print " @folders\n";
	my $inbox = $imap->status($folder);
    	print "Messaggi: " . $inbox->{MESSAGES} . "\n";
	print Data::Dumper::Dumper($inbox);
exit;
	#
        # fetch ID-s that match criteria, sorted by subject and reverse date
	#
        my $messages	= $imap->search({
              FROM   	=> 'noreply@autodesk.com',
              SUBJECT	=> "Your product (Simplicity",
           });
	print "Messaggi trovati: $#{$messages} \n";

        # fetch message summaries (actually, a lot more)
        my $summaries = $imap->get_summaries([@{$messages} ]);
	
	my %record = ();
        foreach (@$summaries) {
		if	($_->uid < 600){next;} # Limite provvisorio

		if 	($_->subject =~ /purchased from/ ) {
			$record{'BUY'} = "ppp";
		}
		elsif 	($_->subject =~ /downloaded from/ ) {
			$record{'BUY'} = "ddd";
		}
		$record{'DATE'} = $_->uid;

	   	print join(', ',$_->uid, $_->date, $record{'BUY'}) . "\n"; # etc.
        }


$imap->logout;


