if exists (select 1 from sysviews where viewname = 'isumBranShip' and vcreator = 'dba') then
	drop view isumBranShip;
end if;

create view isumBranShip (numorder, nomnom, quant)

as 
-- ����������� ��� ��������� ������������ � ����� ����� ������� ��� �������� � ���� ������
-- ����� ����� �-�� �� ����������� ������������ ����� ������,
-- ���������� ��� ��������� ���-�� - � ���������������� �������� (��)

select numorder, nomnom, sum(quant) as quant
from
	itemBranShip
group by
	numorder, nomnom
;
