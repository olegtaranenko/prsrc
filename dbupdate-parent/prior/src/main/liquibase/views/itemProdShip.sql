if exists (select 1 from sysviews where viewname = 'itemProdShip' and vcreator = 'dba') then
	drop view itemProdShip;
end if;



create view itemProdShip (
	numorder
	, nomnom
	, quant
	, date1
)
-- ����������� ������������ ������ �������, ������� �������.
-- �� ������ �������� ��� �� ������ ��������� ������������ ������ ����� ������� - ���� �������
as 

select 
	  numorder
	, nomnom
	, round(fn.quantity * io.quant, 5) as quant
	, io.outdate
-- ������� ������������� ����� ����������� �������, � ����� ������������ ������������� �������.
from 
	xpredmetybyizdeliaout io 
join 
	itemWareFixe fn 
		on io.prid = fn.productid

	union all
select 
	  io.numorder
	, v.nomnom
	, round(p.quantity * io.quant, 5) as quant
	, io.outdate
-- ������� ���������� ����� ����������� �������
from 
	xpredmetybyizdeliaout io 
join 
	xvariantnomenc v 
		on v.numorder = io.numorder and v.prid = io.prid and v.prext = io.prext
join 
	sproducts p 
		on p.productid = io.prid and v.nomnom = p.nomnom

    union all
select 
	  io.numorder
	, io.nomnom
	, io.quant
	, io.outdate
from 
	xpredmetybynomenkout io
;
