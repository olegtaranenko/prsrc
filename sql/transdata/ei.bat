@echo off

call isql ei_cor_prior.sql prior
rem call isql ei_cor_comtex.sql pm
rem call isql ei_cor_comtex.sql mm
call isql ei_cor_comtex.sql stime
call transdata
