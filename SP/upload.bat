@echo off
if (%1) == () goto usage

set patch=%1

mkdir %patch%
cd %patch%
xcopy \\tsclient\C\dev\mainline\SP\%patch%\*.rar . /Y /S /D
echo | time | date > \\tsclient\C\dev\mainline\SP\%patch%\uploaded
rar x %patch% ..

goto done


:usage
echo Usage %~nx0  patch_name

:done

