ALTER VIEW "DBA"."isumBranShip" (numorder, nomnom, quant)

as 
-- ����������� ��� ��������� ������������ � ����� ����� ������� ��� �������� � ���� ������
-- ����� ����� �-�� �� ����������� ������������ ����� ������,
-- ���������� ��� ��������� ���-�� - � ���������������� �������� (��)

select numorder, nomnom, sum(quant) as quant
from
	itemBranShip
group by
	numorder, nomnom
