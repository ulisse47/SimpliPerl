#!/usr/bin/perl
use strict;
use warnings;
use File::Basename;
use File::DirWalk;
use PDF::Reuse;

my $Path = "E:\\FileServer\\Ricevute\\";
#my $Path = "C:\\Tmp\\";

my ($sec, $min, $hour, $mday, $mon, $year, $wday, $yday, $isdst) = localtime(time);
$year += 1900;  $mon	+= 1;
my $tempo =   join('/', $mday,$mon,$year) . " " .  join(':', $hour,$min,$sec) ;

open(LISTA,">>" . $Path . "Log\\Tutti.log");
print LISTA "+++ $tempo Inizio elaborazione +++\n*** Lista segnalazioni\n";

#my ($d, $dir, @Files);
my $dw = new File::DirWalk;
$dw->onDirEnter(sub {
	my ($dir) = @_;
	my $len = $dw->getDepth();
	if (($dir  =~ /.*Email$/ or $dir =~ /Stampa$/) and ($len = 7))  {

		#unlink 	"$dir\\_tutti.pdf";

		if (stat "$dir\\_tutti.pdf" ) {
			#print "$dir\\_tutti.pdf -->  Esiste\n";
		}else {
			my @Files = <$dir\\*.pdf>;
			#print @Files;
			if ( @Files) {
				prDocDir("$dir");
				prFile("_tutti.pdf");
				prDoc($_) for @Files ;
				prEnd();
				undef @Files;
				print LISTA "$dir\\_tutti.pdf -->  Creato ","\n";
			}
		}

	}

        return File::DirWalk::SUCCESS;
 });
$dw->setDepth(6);
$dw->walk($Path);

print LISTA "+++ $tempo Fine elaborazione +++\n\n";
