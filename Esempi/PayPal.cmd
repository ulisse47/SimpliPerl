echo off
t:
cd /bin
perl -w nPaypal.pl -p -n >>T:\log\PayPalCmd.log
vim T:\log\PayPalCmd.log
