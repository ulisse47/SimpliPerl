#!/usr/bin/perl -w
#	
#       
#	

use strict;
use DBI;
#
# Definizione variabili globali
#
my $database ='l11ustn9_spalleponti'; 
my $hostname = "lhcp1054.webapps.net"; 
my $port = "3306"; 
my $user = 'l11ustn9_admin';
my $password = 'ZhB}@@,oJ_t?';

my $dsn = "DBI:mysql:database=$database;host=$hostname;port=$port"; 

#$dbh = DBI->connect($dsn, $user, $password); 


my $dbh = DBI->connect($dsn, $user, $password, {RaiseError => 1}); 
my $drh = DBI->install_driver("mysql"); 

my $sth = $dbh->prepare("SELECT * FROM status"); 
$sth->execute; 
#
#  Impostazione ambiente
#

