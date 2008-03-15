@echo off
if (%1) == () goto usage
set patch=%1

rar a %patch%\%patch% %patch%\*.* -r -x*.rar
for %%i in ( %patch%\*.sql) do call spdelink %%i
goto done

:usage
echo %~nx0 [patchNumber]
echo WARNING: не использовать для ddl и dml файлы с расширением sql
echo 	или которые не есть ЛИНК на соответствующий файл в директории transdata.
goto done
:done


