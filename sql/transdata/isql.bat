@echo off
if "%1"=="" goto usage
set SQL_FILE=%1

shift
set ALIAS_DSN=%1
if "%ALIAS_DSN%"=="" goto ext_dsn

if /i not %ALIAS_DSN%==prior goto check_stime 
set DSN=%DSN_PRIOR%
if "%DSN%"=="" set DSN=DSN=prior;UID=dba;PWD=sql 
goto check_isql_home

:check_stime
if /i not %ALIAS_DSN%==stime goto check_pm
set DSN=%DSN_STIME%
if "%DSN%"=="" set DSN=DSN=stime;UID=admin;PWD=z 
goto check_isql_home

:check_pm
if /i not %ALIAS_DSN%==pm goto check_mm 
set DSN=%DSN_PM%
if "%DSN%"=="" set DSN=DSN=accountN;UID=admin;PWD=z 
goto check_isql_home

:check_mm
if /i not %ALIAS_DSN%==mm goto ext_dsn
set DSN=%DSN_MM%
if "%DSN%"=="" set DSN=DSN=markmaster;UID=admin;PWD=z 
goto check_isql_home

:ext_dsn
if "%DSN%"=="" goto badenvir
:check_isql_home
set ISQL_HOME=%asa8bin%
if "%ISQL_HOME%"=="" set ISQL_HOME=C:\Program Files\Common Files\Comtec Shared\Sql8
if "%ISQL_HOME%"=="" goto badenvir

:exec
echo Execute file "%SQL_FILE%" on "%DSN%
"%ISQL_HOME%\dbisqlc.exe" -c %DSN% -q %SQL_FILE%
goto done

:usage
echo **********************************************
echo * Usage: %~nx0 script [alias_dsn]
echo *
echo * Sensitive environment variables:
echo *	DSN=%DSN%
echo *	ISQL_HOME=%ISQL_HOME%
echo **********************************************
goto done


:badenvir
echo You must specify environment variables DSN and ASA8BIN
echo For example:
echo	set DSN=DSN=_1;UID=admin;PWD=z
echo	asa8bin=C:\Program Files\Common Files\Comtec Shared\Sql8
echo NOTE!
echo	You should not enclose variables in double quotas
goto done


:done