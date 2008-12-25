@echo off
if (%1) == () goto usage
set patch=%1
mkdir dump\%patch%
cd dump\%patch%
mkdir pm
mkdir mm
mkdir stime
mkdir prior

dbbackup -c "DSN=prior;UID=dba;PWD=sql" -x prior
dbbackup -c "DSN=accountn;UID=admin;PWD=z" -x pm
dbbackup -c "DSN=markmaster;UID=admin;PWD=z" -x mm
dbbackup -c "DSN=stime;UID=admin;PWD=z" -x stime

rar a -r %patch%
if (%2) == (noload) goto done

:: publish to the ftp-server
echo open ftp.petmas.ru>ftpcmd.txt
echo a>>ftpcmd.txt
echo user admin RovWaig4>>ftpcmd.txt
echo cd /pub/taranenko/dump>>ftpcmd.txt
::echo lcd %patch%>>ftpcmd.txt
echo binary>>ftpcmd.txt
echo send %patch%.rar>>ftpcmd.txt
echo bye>>ftpcmd.txt
ftp -s:ftpcmd.txt
cd ..
rd %patch% /q/s
cd ..
goto done

:usage
echo usage: %~nx0 filename [noload]
echo 	noload option skip transfer to remote computer.
:done