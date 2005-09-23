begin
declare icount int;
declare folder_id integer;
declare v_id_cur integer;

select count(*) into icount from size;
if icount <= 0 then
	call legacy_guides();
end if;


--Производственные склады
select id into folder_id from voc_names_st where nm = 'Склады' and belong_id = 0;
call move_old_voc_names(folder_id);
call legacy_sklad();


--Объекты затрат
select id into folder_id from voc_names_st where nm = 'Объекты затрат' and belong_id = 0;
call move_old_voc_names(folder_id);
call legacy_zatr();


--Сторонние организации
select id into folder_id from voc_names_st where nm = 'Сторонние организации' and belong_id = 0;
call move_old_voc_names(folder_id);
call legacy_firms();

-- Загрузить валюту "Условная единиа" и курс ее
call legacy_currency();


-- Номенклатура
select id into folder_id from inv_st where nm = 'Материалы' and belong_id = 0;
call move_old_inv(folder_id);
select id into folder_id from inv_st where nm = 'Изделия' and belong_id = 0;
call move_old_inv(folder_id);
call legacy_inv();


-- Заполнение вспомогательных таблиц для вариантных изделий
call host_legacy_variant();

-- Загрузить приходный ордера на склады в аналитическую базу st
call legacy_income_order();


update system set trans_date = now();

commit;

end;
