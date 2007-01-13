@echo off
if (%1) == () goto usage

set patch=%1

mkdir %patch%
cd %patch%
xcopy \\tsclient\C\dev\mainline\SP\%patch%\*.* . /Y
echo | time | date > \\tsclient\C\dev\mainline\SP\%patch%\uploaded


goto done


:usage
echo Usage %~nx0  patch_name

:done

