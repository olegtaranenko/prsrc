@echo off
::echo on

setlocal

::set REP_LOCATION=\\Dima\d\devNext\sql\transdata
::set REP_LOCATION=F:\sql\transdata

if (%1) == () goto usage
set patch=%1
shift

set next_step=%1
shift
if (%next_step%) == () set next_step=code

:loop
if /i (%next_step%) == (ddl) goto start_ddl
if /i (%next_step%) == (all) goto start_ddl
goto code
:start_ddl
if (%ddl_executed%) == (1) goto code
set ddl_executed=1
echo Executing DDL...
if not exist %patch%\ddl.lst set missed_file=ddl.lst&goto missed_warn
for /F "tokens=*" %%i in (%patch%\ddl.lst) do call ../sql/transdata/isql.bat %patch%\%%i 

:code
if /i (%next_step%) == (code) goto start_code
if /i (%next_step%) == (all) goto start_code
goto dml
:start_code
if (%code_executed%) == (1) goto dml
echo Executing CODE...
set code_executed=1
if exist %patch%\logging_common.sql	call ../sql/transdata/isql.bat %patch%\logging_common.sql	prior	
if exist %patch%\logging_common.sql	call ../sql/transdata/isql.bat %patch%\logging_common.sql	stime
if exist %patch%\logging_common.sql	call ../sql/transdata/isql.bat %patch%\logging_common.sql	pm
if exist %patch%\logging_common.sql	call ../sql/transdata/isql.bat %patch%\logging_common.sql	mm

if exist %patch%\slave_prior.sql	call ../sql/transdata/isql.bat %patch%\slave_prior.sql	prior
if exist %patch%\slave_common.sql	call ../sql/transdata/isql.bat %patch%\slave_common.sql	prior
if exist %patch%\slave_common.sql	call ../sql/transdata/isql.bat %patch%\slave_common.sql	stime
if exist %patch%\slave_common.sql	call ../sql/transdata/isql.bat %patch%\slave_common.sql	pm
if exist %patch%\slave_common.sql	call ../sql/transdata/isql.bat %patch%\slave_common.sql	mm
if exist %patch%\slave_comtex.sql	call ../sql/transdata/isql.bat %patch%\slave_comtex.sql	stime
if exist %patch%\slave_comtex.sql	call ../sql/transdata/isql.bat %patch%\slave_comtex.sql	pm
if exist %patch%\slave_comtex.sql	call ../sql/transdata/isql.bat %patch%\slave_comtex.sql	mm
if exist %patch%\slave_stime.sql	call ../sql/transdata/isql.bat %patch%\slave_stime.sql	stime

if exist %patch%\build_host_proc.sql	call ../sql/transdata/isql.bat %patch%\build_host_proc.sql	prior
if exist %patch%\remote_prior.sql	call ../sql/transdata/isql.bat %patch%\remote_prior.sql	prior
if exist %patch%\build_host_proc.sql	call ../sql/transdata/isql.bat %patch%\build_host_proc.sql	stime
if exist %patch%\remote_comtex.sql	call ../sql/transdata/isql.bat %patch%\remote_comtex.sql	stime
if exist %patch%\remote_stime.sql	call ../sql/transdata/isql.bat %patch%\remote_stime.sql	stime
if exist %patch%\build_host_proc.sql	call ../sql/transdata/isql.bat %patch%\build_host_proc.sql	pm
if exist %patch%\remote_comtex.sql	call ../sql/transdata/isql.bat %patch%\remote_comtex.sql	pm
if exist %patch%\build_host_proc.sql	call ../sql/transdata/isql.bat %patch%\build_host_proc.sql	mm
if exist %patch%\remote_comtex.sql	call ../sql/transdata/isql.bat %patch%\remote_comtex.sql	mm

if exist %patch%\server_common.sql	call ../sql/transdata/isql.bat %patch%\server_common.sql	prior
if exist %patch%\server_common.sql	call ../sql/transdata/isql.bat %patch%\server_common.sql	stime
if exist %patch%\server_common.sql	call ../sql/transdata/isql.bat %patch%\server_common.sql	pm
if exist %patch%\server_common.sql	call ../sql/transdata/isql.bat %patch%\server_common.sql	mm
if exist %patch%\server_prior.sql	call ../sql/transdata/isql.bat %patch%\server_prior.sql	prior
if exist %patch%\server_comtex.sql	call ../sql/transdata/isql.bat %patch%\server_comtex.sql	stime
if exist %patch%\server_comtex.sql	call ../sql/transdata/isql.bat %patch%\server_comtex.sql	pm
if exist %patch%\server_comtex.sql	call ../sql/transdata/isql.bat %patch%\server_comtex.sql	mm

if exist %patch%\servertables_prior.sql	call ../sql/transdata/isql.bat %patch%\servertables_prior.sql	prior

if exist %patch%\codebase_prior.sql	call ../sql/transdata/isql.bat %patch%\codebase_prior.sql	prior
if exist %patch%\prior_views.sql	call ../sql/transdata/isql.bat %patch%\prior_views.sql	prior
if exist %patch%\codebase_comtex.sql	call ../sql/transdata/isql.bat %patch%\codebase_comtex.sql	stime
if exist %patch%\codebase_comtex.sql	call ../sql/transdata/isql.bat %patch%\codebase_comtex.sql	pm
if exist %patch%\codebase_comtex.sql	call ../sql/transdata/isql.bat %patch%\codebase_comtex.sql	mm
if exist %patch%\codebase_stime.sql	call ../sql/transdata/isql.bat %patch%\codebase_stime.sql	stime

if exist %patch%\runonce_after_prior.sql	call ../sql/transdata/isql.bat %patch%\runonce_after_prior.sql	prior
if exist %patch%\runonce_after_mm.sql	call ../sql/transdata/isql.bat %patch%\runonce_after_mm.sql	mm
::xcopy  %patch%\*.sql %REP_LOCATION% /Y

if exist %patch%\runonce_after.bat	call %patch%\runonce_after.bat


:dml
if /i (%next_step%) == (dml) goto start_dml
if /i (%next_step%) == (all) goto start_dml
goto list_file
:start_dml
if (%dml_executed%) == (1) goto list_file
set dml_executed=1
echo Executing DML...
if not exist %patch%\dml.lst set missed_file=dml.lst&goto missed_warn
for /F "tokens=*" %%i in (%patch%\dml.lst) do call ../sql/transdata/isql.bat %patch%\%%i 
::for /F "tokens=*" %%i in (%patch%\sql.list) do echo call ../sql/transdata/isql.bat %patch%\%%i & xcopy  %patch%\%%i \\Dima\d\devNext\sql\transdata

:list_file
::if not exist (%patch%\%next_step%.lst) goto next_done
if /i (%next_step%) == (ddl) goto next_done
if /i (%next_step%) == (code) goto next_done
if /i (%next_step%) == (dml) goto next_done
if /i (%next_step%) == (all) goto next_done
echo Executing %next_step%.lst...
for /F "tokens=*" %%i in (%patch%\%next_step%.lst) do call ../sql/transdata/isql.bat %patch%\%%i 
goto next_done

:next_done
set next_step=%1
shift
if (%next_step%) == () goto done
goto loop

:missed_warn
if /i not (%next_step%) == (all) echo file %patch%\%missed_file% does not exists
goto loop


:usage
echo Usage %~nx0 patch_name {options}
echo 	Opitons: ddl, code, dml, 

:done

endlocal

