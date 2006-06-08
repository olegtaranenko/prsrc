-- Удалить все remote таблицы
-- Впредь будем использовать только процедуры!
call build_remote_table('inv', 0);
call build_remote_table('jmat', 0);
call build_remote_table('mat', 0);
call build_remote_table('jscet', 0);
call build_remote_table('scet', 0);

-- не удалять 
--call build_remote_table('voc_names', 0);
