echo off
t:
cd /bin
perl -w SimpDown.pl -u -m >T:\log\SimpDown.log
gvim T:\log\SimpDown.log
pause
