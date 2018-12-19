echo off
t:
cd /bin
perl -w ImapPaypal.pl -n -u -p -e -m >>T:\log\PayPalCmd.log
vim T:\log\PayPalCmd.log
