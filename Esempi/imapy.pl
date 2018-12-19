           use Net::IMAP::Client;

           my $imap = Net::IMAP::Client->new(

               server => 'mail.you.com',
               user   => 'USERID',
               pass   => 'PASSWORD',
               ssl    => 1,                              # (use SSL? default no)
               ssl_verify_peer => 1,                     # (use ca to verify server, default yes)
               ssl_ca_file => '/etc/ssl/certs/certa.pm', # (CA file used for verify server) or
             # ssl_ca_path => '/etc/ssl/certs/',         # (CA path used for SSL)
               port   => 993                             # (but defaults are sane)

           ) or die "Could not connect to IMAP server";

           # everything's useless if you can't login
           $imap->login or
             die('Login failed: ' . $imap->last_error);

           # let's see what this server knows (result cached on first call)
           my $capab = $imap->capability;
              # or
           my $knows_sort = $imap->capability( qr/^sort/i );

           # get list of folders
           my @folders = $imap->folders;

           # get total # of messages, # of unseen messages etc. (fast!)
           my $status = $imap->status(@folders); # hash ref!

           # select folder
           $imap->select('INBOX');

           # get folder hierarchy separator (cached at first call)
           my $sep = $imap->separator;

           # fetch all message ids (as array reference)
           my $messages = $imap->search('ALL');

           # fetch all ID-s sorted by subject
           my $messages = $imap->search('ALL', 'SUBJECT');
              # or
           my $messages = $imap->search('ALL', [ 'SUBJECT' ]);

           # fetch ID-s that match criteria, sorted by subject and reverse date
           my $messages = $imap->search({
               FROM    => 'foo',
               SUBJECT => 'bar',
           }, [ 'SUBJECT', '^DATE' ]);

           # fetch message summaries (actually, a lot more)
           my $summaries = $imap->get_summaries([ @msg_ids ]);

           foreach (@$summaries) {
               print $_->uid, $_->subject, $_->date, $_->rfc822_size;
               print join(', ', @{$_->from}); # etc.
           }

           # fetch full message
           my $data = $imap->get_rfc822_body($msg_id);
           print $$data; # it's reference to a scalar

           # fetch full messages
           my @msgs = $imap->get_rfc822_body([ @msg_ids ]);
           print $$_ for (@msgs);

           # fetch single attachment (message part)
           my $data = $imap->get_part_body($msg_id, '1.2');

           # fetch multiple attachments at once
           my $hash = $imap->get_parts_bodies($msg_id, [ '1.2', '1.3', '2.2' ]);
           my $part1_2 = $hash->{'1.2'};
           my $part1_3 = $hash->{'1.3'};
           my $part2_2 = $hash->{'2.2'};
           print $$part1_2;              # need to dereference it


