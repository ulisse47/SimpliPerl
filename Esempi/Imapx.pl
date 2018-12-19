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

print "ok";
# Use the module

use Net::IMAP::Client;;
use Data::Dumper;
use Email::MIME::Encodings;

print "ok";
	my $imap = Net::IMAP::Client->new(
	   	server	 => 'pop.ccprogetti.it',
       		username => 'assistenza@ccprogetti.it',
      		password => 'ccprogetti',	   
		port	 => 143
		) or die "Could not connect to IMAP server";
	

	$imap->login( 'assistenza@ccprogetti.it','ccprogetti') or
		die('Login failed: ');


	my $folder = 'INBOX.Autodesk App Store';

	my $nm = $imap->examine($folder) or 
		die "Could not select: " . $@ . "\n";

print "ok";        # get list of folders
        my @folders = $imap->folders;
	print " @folders\n";
#
exit;
#
	my $inbox = $imap->status($folder);
    	print "Messaggi non letti: " . $inbox->{UNSEEN} . "\n";
    	print Data::Dumper::Dumper($inbox);

        # fetch ID-s that match criteria, sorted by subject and reverse date
        my $messages = $imap->search({
               FROM    => 'noreply@autodesk.com',
#              SUBJECT => 'bar',
           }, [ 'SUBJECT', '^DATE' ]);
   	 
	print " @{$messages}[1] \n";

           # fetch message summaries (actually, a lot more)
           my $summaries = $imap->get_summaries([@{$messages} ]);

           foreach (@$summaries) {
               print $_->uid, $_->subject, $_->date, $_->rfc822_size;
               print join(', ', @{$_->from} . "\n"); # etc.
           }

           # fetch multiple attachments at once
#           my $hash = $imap->get_parts_bodies($messages, [ '1.2', '1.3', '2.2' ]);
#           my $part1_2 = $hash->{'1.2'};
#           my $part1_3 = $hash->{'1.3'};
#           my $part2_2 = $hash->{'2.2'};
#           print $$part1_2;              # need to dereference it
	 
        # fetch full messages
#        my @msgs = $imap->get_rfc822_body([  1,2,3 ]);
#        print $$_ for (@msgs);
	 
        # fetch full message
#        my $data = $imap->get_rfc822_body(1);
#        print $$data; # it's reference to a scalar

    my $summary = $imap->get_summaries(10)->[0];
#    my $part = $summaries->get_subpart('1.1');
    my $body = $imap->get_rfc822_body([1,2]);
    my $cte = $body->transfer_encoding;  # Content-Transfer-Encoding
       $body = Email::MIME::Encodings::decode($cte, $$body);	 
	print $body;


$imap->logout;


