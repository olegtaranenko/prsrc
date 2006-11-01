
-- Изменения в id-образовании Комтеха, которые пришли в версии
-- Комтех 8.1.5 (январь 2006 года).
-- Появилась таблица inc_table, в которую похоже складывается 
-- номер id, для таблицы который должен быть использован для 
-- добавления следующей записи в таблицу.
-- Кроме этих изменений, которые просто исправят текущую ситуацию
-- потребуется еще внести изменения в процедуры типа insert_host, 
-- insert_remote и, может быть, других для корректного исправления
-- этих id.
-- TODO!!! Можно будет также изменить с целью увеличения 
-- производительности процедуру получения следующего глобального id 
-- для комтеховской таблицы

-- 31.10.2006 
--	Из-за бага: Ошибка при смене Предприяятия в приходных накладныч
--	потребовалось изменить логику получения/фиксирования nextid. Теперь она
--	действительно берется из inc_table.

begin
	declare nxt_id integer;

   	for d_cur as dc dynamic scroll cursor for
   		select table_nm as r_table_nm from inc_table
   	for update
	do
		execute immediate 'select max(id) into nxt_id from ' + r_table_nm;
		set nxt_id = isnull(nxt_id, 0) + 1;
		update inc_table set next_id = nxt_id where current of dc;

		call build_id_track_trigger(r_table_nm);
	end for;
end;



commit;
