if exists (select 1 from sysviews where viewname = 'BayNomenk') then
	drop view BayNomenk
end if;
 

CREATE VIEW
	BayNomenk
AS
SELECT
	 numOrder
	,nomNom
	,quant
	,cenaEd
	,id_scet
from 
	xPredmetyByNomenk 
