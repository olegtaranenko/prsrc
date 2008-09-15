::@echo off
if (%1) == () goto usage
set patch=%1& shift

:: delink junction files
rar a %patch%\%patch% %patch%\*.* -r -x*.rar
for %%i in ( %patch%\*.sql) do call spdelink %%i

:: retrieve dump from the ftp-server
echo open www.markmaster.ru>ftpcmd.txt
echo markmaster>>ftpcmd.txt
echo ItUs3P@ss>>ftpcmd.txt
echo cd /pub/taranenko/sp>>ftpcmd.txt
echo lcd %patch%>>ftpcmd.txt
echo binary>>ftpcmd.txt
echo send %patch%.rar>>ftpcmd.txt
echo bye>>ftpcmd.txt
ftp -s:ftpcmd.txt
goto done

:usage
echo %~nx0 [patchNumber]
echo WARNING: не использовать для ddl и dml файлы с расширением sql
echo 	или которые не есть ЛИНК на соответствующий файл в директории transdata.
goto done
:done
