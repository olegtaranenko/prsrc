exit;

--������� ����� ��������� � "���� �������"
update bayorders set firmId = f.firmId
from bayguidefirms f where f.name = '���� �������'
and bayorders.numorder = 9051521;

update bayorders set id_bill = null where numorder = 9051521;
update bayorders set id_bill = 11010 where numorder = 9051521;

delete from bayguidefirms where name = '���� �������_1';

delete from bayorders where numorder = 9051513;

-- ��������� � ������ ����� ������ ��������
update bayguidefirms a
set id_voc_names = f.id_voc_names
from bayguidefirms f 
where f.name = '������_1'
and a.name = '������';

update bayguidefirms 
set id_voc_names = null
where name = '������_1';

update bayguidefirms 
set name = '��� ������'
where name = '������';


update bayorders set firmId = f.firmId
from bayguidefirms f 
where f.name = '��� ������'
and bayorders.numorder = 9051523;

delete from bayorders where numorder = 9051503;


delete from bayguidefirms where name = '������_1';
