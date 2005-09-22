@echo off

call isql logging_common prior
call isql logging_common stime
call isql logging_common pm
call isql logging_common mm

call isql slave_prior.sql prior
call isql prior_db_change.sql prior

call isql slave_common.sql prior
call isql slave_common.sql stime
call isql slave_common.sql pm
call isql slave_common.sql mm

call isql slave_comtex.sql stime
call isql slave_comtex.sql pm
call isql slave_comtex.sql mm


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

call isql codebase_prior.sql prior

call isql codebase_comtex.sql stime
call isql codebase_comtex.sql pm
call isql codebase_comtex.sql mm
call isql codebase_stime.sql stime



call isql legacy_buh.sql stime
call isql legacy_buh.sql pm
call isql legacy_buh.sql mm
call isql legacy_prior.sql prior

call isql legacy_driver.sql prior

call isql drop_remote_prior.sql prior

rem call isql forbid_prior.sql prior

rem exit
