ALTER VIEW "DBA"."itemSellShip" (numorder, nomnom, quant, date1)
-- ����������� ������������ ������ �������, ������� �������.
-- �� ������ �������� ��� �� ������ ��������� ������������ ������ ����� ������� - ���� �������
as 
select io.numorder, io.nomnom, io.quant * n.perlist, outdate
from baynomenkout io
join sguidenomenk n on n.nomnom = io.nomnom
