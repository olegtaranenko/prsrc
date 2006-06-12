@echo off

call isql logging_common prior
call isql logging_common stime
call isql logging_common pm
call isql logging_common mm

rem call isql dbcc_prior.sql prior

call isql slave_prior.sql prior
rem call isql prior_db_change.sql prior

call isql slave_common.sql prior
call isql slave_common.sql stime
call isql slave_common.sql pm
call isql slave_common.sql mm

call isql slave_comtex.sql stime
call isql slave_comtex.sql pm
call isql slave_comtex.sql mm

call isql slave_stime.sql stime


call isql build_host_proc.sql prior
call isql remote_prior.sql prior
call isql build_host_proc.sql stime
call isql remote_comtex.sql   stime
call isql remote_stime.sql    stime
call isql build_host_proc.sql pm
call isql remote_comtex.sql   pm
call isql build_host_proc.sql mm
call isql remote_comtex.sql   mm

call isql server_common.sql prior
call isql server_common.sql stime
call isql server_common.sql pm
call isql server_common.sql mm

call isql server_prior.sql prior
call isql server_comtex.sql stime
call isql server_comtex.sql pm
call isql server_comtex.sql mm

call isql servertables_prior.sql prior

rem call isql dml_before_prior.sql prior
::call isql dbcc_prior.sql prior

call isql codebase_prior.sql prior

call isql codebase_comtex.sql stime
call isql codebase_comtex.sql pm
call isql codebase_comtex.sql mm
call isql codebase_stime.sql stime

call isql comtex_db_change stime
call isql comtex_db_change pm
call isql comtex_db_change mm

rem call isql dml_after_prior.sql prior

::call isql dbcc_stime.sql stime
rem call isql dbcc_comtex.sql pm
rem call isql dbcc_comtex.sql mm
rem call isql dbcc_comtex.sql stime

rem call isql legacy_buh.sql stime
rem call isql legacy_buh.sql pm
rem call isql legacy_buh.sql mm
rem call isql legacy_prior.sql prior

rem call isql legacy_driver.sql prior

rem call isql drop_remote_prior.sql prior

rem call isql fill_venture_order.sql prior

rem exit
