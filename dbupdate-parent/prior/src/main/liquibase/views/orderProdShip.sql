if exists (select 1 from sysviews where viewname = 'orderProdShip' and vcreator = 'dba') then
	drop view orderProdShip;
end if;



create view orderProdShip (
	numorder
	, quant
	, date1
)
-- ����������� ������������ ������ �������, ������� �������.
-- �� ������ �������� ��� �� ������ ��������� ������������ ������ ����� ������� - ���� �������
as 

select 
