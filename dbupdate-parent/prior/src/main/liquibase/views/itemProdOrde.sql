if exists (select 1 from sysviews where viewname = 'itemProdOrde' and vcreator = 'dba') then
	drop view itemProdOrde;
end if;

create view itemProdOrde (numorder, nomnom, quant, cenaEd, prid, prext, quantEd)
-- ������ ������������, �������� � �������� ������ ������������
-- ����� ������� � rowmat ������ �� �������� ������!!!
-- ������� (������� ����������) ����������� �� ��������� ������������.
-- quantEd - ���������� ��������� ��������� ������������ � ���� �������
-- � ���������������� ��������!
as 
select numorder, nomnom, round(fn.quantity * io.quant, 5), null, io.prid, io.prext, fn.quantity
from xPredmetyByIzdelia io 
join itemWareFixe fn on io.prid = fn.productid
	union all
select io.numorder, v.nomnom, round(p.quantity * io.quant, 5), null, io.prid, io.prext, p.quantity
from xPredmetyByIzdelia io
join xVariantNomenc v on v.numorder = io.numorder and v.prid = io.prid and v.prext = io.prext
join sProducts p on p.productid = io.prid and v.nomnom = p.nomnom
	union all
select po.numorder, po.nomnom, po.quant, po.cenaEd, null, null, 1
from xPredmetyByNomenk po;