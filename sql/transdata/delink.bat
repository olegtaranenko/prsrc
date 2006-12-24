@echo off


:loop
set cur_file=%1
set file_path=%~dp1
shift
if (%cur_file%) == () goto done
if /i %file_path% == %PM_DEV_ROOT%\sql\transdata\ goto panic
echo %file_path%\%cur_file%

if exist %temp%\%cur_file%.bak del %cur_file%.bak
copy /b %cur_file% %temp%\%cur_file%.bak
del %cur_file%
copy /b /y %PM_DEV_ROOT%\sql\transdata\%cur_file% %file_path%
goto loop

:panic
echo Panic! DO NOT DELETE ORIGINAL FILE!
exit

:done

