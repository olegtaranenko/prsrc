--=====================================================
--	Ship ���������� �� Ship: ��������� �� ������������
--=====================================================

if exists (select 1 from sysviews where viewname = 'itemSellShip' and vcreator = 'dba') then
	drop view itemSellShip;
end if;

create view itemSellShip (numorder, nomnom, quant, date1)
-- ����������� ������������ ������ �������, ������� �������.
-- �� ������ �������� ��� �� ������ ��������� ������������ ������ ����� ������� - ���� �������
as 
select io.numorder, io.nomnom, io.quant * n.perlist, outdate
from baynomenkout io
join sguidenomenk n on n.nomnom = io.nomnom
;
