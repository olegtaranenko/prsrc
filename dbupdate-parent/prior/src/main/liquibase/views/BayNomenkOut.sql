if exists (select 1 from sysviews where viewname = 'BayNomenkOut') then
	drop view BayNomenkOut
end if;
 

CREATE VIEW
	BayNomenkOut
AS
SELECT
	 outDate
	,numOrder
	,nomNom
	,quant
	,id_mat
	,id_jmat
from 
	xPredmetyByNomenkOut 
