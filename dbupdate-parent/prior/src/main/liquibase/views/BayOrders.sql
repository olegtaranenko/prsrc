if exists (select 1 from sysviews where viewname = 'BayOrders') then
	drop view BayOrders
end if;
 

CREATE VIEW
	BayOrders
AS
SELECT
	numOrder
	,inDate
	,ManagId
	,StatusId
	,ProblemId
	,FirmId
	,outDateTime
	,ordered
	,paid
	,shipped
	,lastManagId
	,Invoice
	,Remark
	,Transport
	,id_jscet
	,ventureid
	,id_bill
	,rate
from 
	Orders 
where 
	WerkId = 1