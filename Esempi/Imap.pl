#!perl -w Imap.pl
#Imap.pl
#
#	Legge la posta IMAP nella cartella 
#	
#	
#
# Always be safe
use strict;
#use warnings;

# Use the module
use Mail::IMAPClient;#print "hello";

	my $imaps = Mail::IMAPClient->new( 	Server  => 'pop.ccprogetti.it:110',
                                		User    => 'cavalierisrl@ccprogetti.it',
                              			Password=> 'Ulisse@1947') or die "IMAP Failure: $@";

my $box;
 	foreach $box qw( HAM SPAM ) %{
  	# Which file are the messages going into
   		my $file = "mail/$box";

   		# Select the mailbox to get messages from
   		$imap->select($box) or die "IMAP Select Error: $!";

   		# Store each message as an array element
  		my @msgs = $imap->search('ALL') or die "Couldn't get all messages\n";

   		# Loop over the messages and store in file
   		foreach my $msg (@msgs) {
     		# Pipe msgs through 'formail' so they are stored properly
     		open my $pipe, "| formail >> $file" or die("Formail Open Pipe Error: $!");

   		}
	# Send msg through file pipe
     	$imap->message_to_file($pipe, $msg);
     	# Close the folder
     	$imap->close($box);
     	}

	# We're all done with IMAP here
	$imaps->logout();

