@echo off
if (%1) == () goto usage

set patch=%1

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
if exist %patch%\codebase_comtex.sql	call ../sql/transdata/isql.bat %patch%\codebase_comtex.sql	stime
if exist %patch%\codebase_comtex.sql	call ../sql/transdata/isql.bat %patch%\codebase_comtex.sql	pm
if exist %patch%\codebase_comtex.sql	call ../sql/transdata/isql.bat %patch%\codebase_comtex.sql	mm
if exist %patch%\codebase_stime.sql	call ../sql/transdata/isql.bat %patch%\codebase_stime.sql	stime

if exist %patch%\runonce_after_prior.sql	call ../sql/transdata/isql.bat %patch%\runonce_after_prior.sql	prior
if exist %patch%\runonce_after_mm.sql	call ../sql/transdata/isql.bat %patch%\runonce_after_mm.sql	mm

if exist %patch%\runonce_after.bat	call %patch%\runonce_after.bat

goto done


:usage
echo Usage %~nx0  patch_name

:done

