echo off
t:
cd /bin
perl -w GenCodeNT.pl -u -m >GenCodeNT.csv
gvim GenCodeNT.csv
pause
