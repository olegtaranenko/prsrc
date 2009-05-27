exit;

--Удалить заказ связанный с "Бюро рекламы"
update bayorders set firmId = f.firmId
from bayguidefirms f where f.name = 'бюро рекламы'
and bayorders.numorder = 9051521;

update bayorders set id_bill = null where numorder = 9051521;
update bayorders set id_bill = 11010 where numorder = 9051521;

delete from bayguidefirms where name = 'бюро рекламы_1';

delete from bayorders where numorder = 9051513;

-- исправить у старой фирмы Фараон название
update bayguidefirms a
set id_voc_names = f.id_voc_names
from bayguidefirms f 
where f.name = 'Фараон_1'
and a.name = 'Фараон';

update bayguidefirms 
set id_voc_names = null
where name = 'Фараон_1';

update bayguidefirms 
set name = 'РПФ Фараон'
where name = 'Фараон';


update bayorders set firmId = f.firmId
from bayguidefirms f 
where f.name = 'РПФ Фараон'
and bayorders.numorder = 9051523;

delete from bayorders where numorder = 9051503;


delete from bayguidefirms where name = 'Фараон_1';
