if exists (select 1 from sysviews where viewname = 'itemWallShip' and vcreator = 'dba') then
	drop view itemWallShip;
end if;


create view itemWallShip (
	  outdate
	, numorder
	, type
	, prId
	, prExt
	, prNomnom
	, cenaEd
	, quant
	, costEd
	, firmName
	, ventureId
	, statusid
	, werkId
)
-- ������ ������� ��������� ����� �� �������� � ���������� �������, ������� ����� ������������� ��� �����������
-- �� ���� ������������ ������������: ���������������� (������� � ��������� ���-��), ������, �������.
-- ������������ � ������ ���������� �������
as 
select 
	  po.outdate
	, po.numorder
	, 1
	, po.prId
	, po.prExt
	, null
	, p.cenaEd
	, po.quant
	, io.costEd
	, f.name
	, isnull(o.ventureId, 3) as ventureId -- ��������� ��� �� ���������
	, o.statusid
	, o.werkId 
from 
	xpredmetybyizdeliaout po
join 
	xpredmetybyizdelia p 
		on p.numorder   = po.numorder 
		and p.prid      = po.prid 
		and p.prext     = po.prext
join 
	orderWareShip io 
		on  io.outdate  = po.outdate 
		and io.numorder = po.numorder 
		and io.prid     = po.prid 
		and io.prext    = po.prext
join 
	orders o 
		on o.numorder   = po.numorder 
		and o.numorder  = p.numorder
join 
	FirmGuide f 
		on f.firmid = o.firmid

	union all 
select 
	  po.outdate
	, po.numorder
	, 2
	, null
	, null
	, po.nomnom
	, p.cenaEd
	, po.quant
	, round(round(n.cost, 2) / n.perlist, 2) as costEd
	, f.name
	, isnull(o.ventureId, 3) as ventureId -- ��������� ��� �� ���������
	, o.statusid
	, o.werkId
from 
	xpredmetybynomenkout po
join 
	xpredmetybynomenk p 
		on p.numorder = po.numorder 
		and p.nomnom = po.nomnom
join 
	sguidenomenk n 
		on n.nomnom = po.nomnom 
		and n.nomnom = p.nomnom
join 
	orders o 
		on o.numorder = po.numorder 
		and o.numorder = p.numorder
join 
	FirmGuide f 
		on f.firmid = o.firmid


	union all 
select 
	  u.outdate
	, u.numorder
	, 4
	, null
	, null
	, null
	, 1.0
	, u.quant
	, 0
	, f.name
	, isnull(o.ventureId, 3) as ventureId -- ��������� ��� �� ���������
	, o.statusid
	, o.werkId
from 
	xuslugout u
join 
	orders o 
		on o.numorder = u.numorder 
join 
	FirmGuide f 
		on f.firmid = o.firmid
