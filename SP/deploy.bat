@echo off
if (%1) == () goto usage
set patch=%1

rar a %patch%\%patch% %patch%\*.* -r -x*.rar
for %%i in ( %patch%\*.sql) do call spdelink %%i
goto done

:usage
echo %~nx0 [patchNumber]
echo WARNING: �� �ᯮ�짮���� ��� ddl � dml 䠩�� � ���७��� sql
echo 	��� ����� �� ���� ���� �� ᮮ⢥�����騩 䠩� � ��४�ਨ transdata.
goto done
:done


