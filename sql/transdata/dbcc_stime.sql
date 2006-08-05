--
-- ����������� ��������� �������� "������������" ����������
-- ��� ��������� �� � ���� ������� ��� ����������� ��� ���������
-- ������� �� ��� ���� ��������
--
/*
-- ������� ������� ������ ������� �� ���� Prior
call build_remote_table('sdocs', 1);


begin
   	for sklad_cur as s dynamic scroll cursor for
		select id as r_id from jmat j
		join sdocs_prior p on p.id_jmat = j.id and p.sourid < -1000 and p.destid < -1000
		where j.id_guide = 1210
	do
		update jmat set id_guide = 1220 
		, tp1=2, tp2=2, tp3=2, tp4= 0 
		, id_s = 3520, id_d = 3519 where id = r_id;
	end for;
end;

-- ������� ������� �� ���������� �������������
call build_remote_table('sdocs', 0);


--update mat set summa = summa_sale where tp1 = 3;

-- ��������� ������������ ����� ������������� �������
-- ��-�� ��� �������� �������� ������.
delete from jmat where id_s = id_d and id > 0;
-- � ������ �/������ ��� ������ ��������� ����������
-- � ���� ����� ������� ��������� �����������.
delete from jmat where dat < '20051013' and id > 0;


-- ��������� ���������
-- ��������� ���������
update mat m set summav = 0, summa_salev = 0 
from jmat j
where 
j.id = m.id_jmat and j.id_guide in (1210, 1120);


-- ����������� ���������
update mat m set summa_sale=0, summav = 0, summa_salev = 0 
from jmat j
where --m.id = 5906 and 
j.id = m.id_jmat and j.id_guide = 1220;

-- �������� �� �����?
update mat m set summa_curr=summa_sale
from jmat j
where 
j.id = m.id_jmat and j.id_guide in ( 1120, 1210 );



update mat m set summa = summa_sale, summa_sale = 0, summa_salev = 0, summav = summa_sale /30 
from jmat j
where 
j.id = m.id_jmat and j.id_guide in ( 1127 );
*/

/*
--**********************************************
--������� � ������� ���� Prior 9 ������� 2006 
--**********************************************


begin 
	declare v_ventureId varchar(20);
	declare v_numdoc varchar(20);

	for mm_income as mm dynamic scroll cursor for
		select id as r_id_jmat, id_code as r_id_analytic from jmat where id_code > 0
	do
		update jmat set id_code = 0 where id = r_id_jmat;
	update jmat set id_code = r_id_analytic where id = r_id_jmat;

	end for;
end;
*/


update mat set summa_sale = summa, summa_salev = summav 
from jmat where mat.id_jmat = jmat.id and jmat.id_guide = 1127


commit;

