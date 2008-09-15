@echo off
if (%1) == () goto usage

set patch=%1

mkdir %patch%
cd %patch%
:: get from the ftp-server
echo open www.markmaster.ru>ftpcmd.txt
echo markmaster>>ftpcmd.txt
echo ItUs3P@ss>>ftpcmd.txt
echo cd /pub/taranenko/sp>>ftpcmd.txt
echo binary>>ftpcmd.txt
echo get %patch%.rar>>ftpcmd.txt
echo bye>>ftpcmd.txt
ftp -s:ftpcmd.txt

goto done


:usage
echo %~nx0 [patch]
goto done
:done
