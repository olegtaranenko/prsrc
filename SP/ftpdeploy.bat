::@echo off
if (%1) == () goto usage
set patch=%1& shift

:: delink junction files
rar a %patch%\%patch% %patch%\*.* -r -x*.rar
for %%i in ( %patch%\*.sql) do call spdelink %%i

:: retrieve dump from the ftp-server
echo open ftp.petmas.ru>ftpcmd.txt
echo a>>ftpcmd.txt
echo user admin RovWaig4>>ftpcmd.txt
echo cd /pub/taranenko/sp>>ftpcmd.txt
echo lcd %patch%>>ftpcmd.txt
echo binary>>ftpcmd.txt
echo send %patch%.rar>>ftpcmd.txt
echo bye>>ftpcmd.txt
ftp -s:ftpcmd.txt
goto done

:usage
echo %~nx0 [patchNumber]
echo WARNING: �� �ᯮ�짮���� ��� ddl � dml 䠩�� � ���७��� sql
echo 	��� ����� �� ���� ���� �� ᮮ⢥�����騩 䠩� � ��४�ਨ transdata.
goto done
:done
