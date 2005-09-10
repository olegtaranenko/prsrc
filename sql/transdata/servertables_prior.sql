-- Создать remote-таблицы. 
-- Они нужны теперь только при загрузке legacy данных

call build_remote_table('voc_names', 1);
call build_remote_table('inv', 1);
call build_remote_table('jmat', 1);
call build_remote_table('mat', 1);
call build_remote_table('jscet', 1);
call build_remote_table('scet', 1);


