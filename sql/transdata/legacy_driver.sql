-- Исправления в фирмах. Зачем-то вместо адреса вписал контакты
-- переносим их в поле rem
begin 
	for aCursor as b1 dynamic scroll cursor for
		select 
			f.firmId as f_id
			,f.name as r_name
			,f.fio as f_fio
			,f.phone as f_phone
			,f.email as f_email
			,id_voc_names as r_id_voc_names
		from guidefirms f
	DO
		call update_host ('voc_names', 'rem', 'address', 'id = ' + convert(varchar(20), r_id_voc_names));
		call update_host ('voc_names', 'address', '''''''''', 'id = ' + convert(varchar(20), r_id_voc_names));
	end for;
end;

begin 
	declare v_rem varchar(100);
	for aCursor as b1 dynamic scroll cursor for
		select 
			f.firmId as f_id
			,f.name as r_name
			,f.fio as f_fio
			,f.phone as f_phone
			,f.email as f_email
			,id_voc_names as r_id_voc_names
		from bayguidefirms f
	DO
		set v_rem = select_remote('stime', 'voc_names', 'rem', 'id = ' + convert(varchar(20), r_id_voc_names));
		if v_rem is null or char_length(v_rem) = 0 then
			call update_host ('voc_names', 'rem', 'address', 'id = ' + convert(varchar(20), r_id_voc_names));
			call update_host ('voc_names', 'address', '''''''''', 'id = ' + convert(varchar(20), r_id_voc_names));
		end if;
	end for;
end;



commit;

/*
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

call legacy_purpose();

update system set trans_date = now();

commit;

end;
*/