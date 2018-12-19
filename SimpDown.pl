#!/usr/bin/perl
use strict;
use Getopt::Long;
use DBI;
use Mail::IMAPClient;
use IO::Socket;
use IO::Socket::SSL;
use DateTime::Format::Mail;
use Data::Dumper;
use Email::MIME::Encodings;
use HTML::Strip;	# Istanza HTML cleaner
	my $hs = HTML::Strip->new();

my %opts = (
            help	=> 0,
	    move	=> 0,
	    update	=> 0
            );

Getopt::Long::Configure('bundling');

GetOptions(
           'h|help'     => \$opts{help},
           'm|move'     => \$opts{move},
           'u|update'   => \$opts{update},
           ) ;

if ($opts{help}) {
   print "\nperl -w $0 [-h|-m|-u]\n 
   	-u : Aggiorna la tabella Simplicity
	-m : Attiva trasferimento messggi 
	-h : Help ";
   exit 0;
}

if ($opts{update})	{ print "\n+++ Attivato aggiornamento DB \n"; }
else			{ print "\n+++ Disattivato aggiornamento DB \n"; }

if ($opts{move})	{ print "\n+++ Attivato trasferimento messaggio \n"; }
else			{ print "\n+++ Disattivato trasferimento messaggio \n"; }

#
# Prepara comandi SQL
#
my $DSN = "dbi:ODBC:CCP";
my $dbh = DBI->connect( $DSN, '', '' ) or die "$DBI::errstr\n";
my $sth = $dbh->prepare( " 
		INSERT INTO Simplicity (
		language, email, fullname, country,
		company, job, data, purchased, product
		) 
		VALUES (?,?,?,?,?,?,?,?,?);
	") or die "$DBI::errstr\n";


# Config  imap connction
my $mail_hostname	= 'pop.ccprogetti.it';
my $mail_hostname_s	= 'pop.securemail.pro';
my $mail_username	= 'assistenza@ccprogetti.it';
my $mail_password	= 'ccprogetti';
my $mail_ssl = 1;

#
# Set folders
#
my $folder	= 'INBOX';
my $arkfolder	= 'INBOX.Autodesk App Store';

# Make sure this is accessable for this namespace
my $socket = undef;

if( $mail_ssl ) {
	# Open up a SSL socket to use with IMAPClient later
	$socket = IO::Socket::SSL->new(
		PeerAddr => $mail_hostname_s,
		PeerPort => 993,
		Timeout => 5,
	);
} else {
	# Open up a none SSL socket to use with IMAPClient later
	$socket = IO::Socket::INET->new(
		PeerAddr => $mail_hostname,
		PeerPort => 143,
		Timeout => 5,
	);
}

# Check we connected
if( ! defined( $socket ) ) {
	print STDERR "Could not open socket to mailserver: $@\n";
	exit 1;
}

my $client = Mail::IMAPClient->new(
	Socket   => $socket,
	User     => $mail_username,
	Password => $mail_password,
	Timeout => 5,
);

# Check we have an imap client
if( ! defined( $client ) ) {
	print STDERR "Could not initialize the imap client: $@\n";
	exit 1;
}

# Check we are authenticated
if( $client->IsAuthenticated() ) {
	# Select the INBOX folder
	if( ! $client->select($folder )) {
		print STDERR "Could not select the INBOX: $@\n";
	} else {
		if( $client->message_count($folder) > 0) {
			print	"\n--- Trovati " . $client->message_count($folder) .
				" messaggi nel folder: <$folder>\n--- account  <$mail_username>\n\n";

			#
			# Elabora tutti i messaggi della cartella
			# We delete messages after processing so get all in the inbox
			#
			my @search = ('SUBJECT "Your product (Simplicity" FROM "Autodesk App Store <noreply@autodesk.com>"');
			my @results = $client->search( @search );
			if (! @results ) {
				print "+-- Nessun messaggio da elaborare\n";
				exit 1;
			} else {print "+-- Inizio elaborazione ...\n";}
			
			my $parse = DateTime::Format::Mail->new( loose => 1 );
			foreach(@results) {
				print "\n+--" . $_ . "\n";

				# Pull the subject out the message
				my $subject = $client->subject( $_ );
				#print $subject ."\n";

				my $product = undef;
				if ( $subject =~ /Your product \((.*)\) has been/ ) {    # parse
            					$product = $1 ;
						#print $product ."\n";
					}
				my $language = undef;
				if ( $subject =~ /from Autodesk App Store <Revit_(.*)>/ ) {    # parse
            					$language = $1 ;
						#print $language ."\n";
					}

				# Pull the body out the message
				my $body = $client->body_string( $_ )
					or die "Could not body_string: ",
					$client->LastError;

				# Try and get the unix time of the message being sent
				# Pull the RFC822 date out the message
				my $date = $client->date( $_ );
				my $mail_date 	= undef;
				my $data_d 	= undef;
				if( $date ) {		
					print $date ."\n";
					$mail_date = $parse->parse_datetime( "$date" );
					#print $mail_date ."\n";
					$data_d = $mail_date->dmy('/') . " " . $mail_date->hms . "\n";
					print $data_d ."\n";
				}

				# Check we have valid stuff
				if( ! $mail_date || ! $subject || ! $body ) {
					print Dumper( $mail_date );
					print Dumper( $subject );
					print Dumper( $body );
					exit 1;
				} else {

					# Decode bodymsg
					my $dcd = undef;
					$dcd = Email::MIME::Encodings::decode(base64 => $body);
					my $c_text = undef;
					$c_text = $hs->parse( $dcd );
					$hs->eof;

					#print Dumper( $c_text );

					my $email = undef;
					if ( $c_text =~ /E-mail: \((.+)\@(.+)\) Full/ ) {    # parse
            					$email = $1 .'@' . $2;
						#print $email ."\n";
					}

					my $fname = undef;
					if ( $c_text =~ /Full Name : (.+) Job/ ) {    # parse
            					$fname = $1 ;
						#print $fname ."\n";
					}

					my $jobtitle = undef;
					if ( $c_text =~ /Job Title : (.+) Company/ ) {    # parse
            					$jobtitle = $1 ;
						#print $jobtitle ."\n";
					}

					my $cname = undef;					
					if ( $c_text =~ /Company Name : (.+) Country/ ) {    # parse
						$cname = $1 ;
						#print $cname ."\n";
					}					

					my $country = undef;					
					if ( $c_text =~ /Country : (.+) Purchased/ ) {    # parse
						$country = $1 ;
						#print $country ."\n";
					}					

					my $purchased = undef;					
					if ( $c_text =~ /Purchased Copies : (\d+) / ) {    # parse
						$purchased = $1 ;
						#print $purchased ."\n";
					}					

					#
					#	Inserisci record in DB
					#
					if ($opts{update}) {
						$sth->execute(
							$language, $email, $fname, $country, 
							$cname, $jobtitle, $data_d, $purchased, $product
						) or die "!!! Errore su inserimento" . $DBI::errstr . "\n";
						print "+++ " .$_ ." registrato su DB \n";
					}

					if ($opts{move}) { # Move the message
						$client->move($arkfolder, $_) 
							or die "Could not move $@\n";
						print "+++ " . $_ ." trasferito alla cartella " . $arkfolder ."\n";
					}
					#print "Msg ok\n";

				} # Fine elaborazione messaggio
			}

			# Delete the messages we have deleted
			# Yes, you read that right, IMAP is strangly awesome
			$client->expunge 
				or die "Could not expunge: $@\n";
			#
			# Fine messaggi da elaborare in INBOX
			#
		} else {
			# No messages
			print "No messages to process\n";
		}
		#
		# Fine elaborazione
		#

		# Close the inbox
		$client->close;
	}
} else {
	print STDERR "Could not authenticate against IMAP: $@\n";
	exit 1;
}

# Tidy up after we are done
$client->done();
$dbh->disconnect or warn $dbh->errstr;
exit 0;;
