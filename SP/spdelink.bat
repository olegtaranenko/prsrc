@echo off

if "%1" == "" goto usage


set cur_filename=%~nx1
set cur_file=%1
set file_path=%~dp1
echo %cur_file%

::goto done
::echo on
::if exist %TEMP%\%cur_filename%.bak del %TEMP%\%cur_filename%.bak
copy /b /y %cur_file% %TEMP%\%cur_filename%.bak
del %cur_file%
copy /b /y ..\sql\transdata\%cur_filename% %cur_file% 
goto done

:panic
echo Panic! DO NOT DELETE ORIGINAL FILE!
goto done

:usage
echo Usage %~nx0 [file1 file2 ...]
echo Run from service pack folder
echo NOTE: Do not forget set variable PM_DEV_ROOT=%PM_DEV_ROOT%
:done

