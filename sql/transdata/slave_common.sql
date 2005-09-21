
--****************************************************************
--                               DELETE
--****************************************************************
if exists (select 1 from sysprocedure where proc_name = 'slave_delete') then
	drop function slave_delete;
end if;

create 
	procedure slave_delete(
		in table_name varchar(50)
		, in where_claus varchar(1000)
		, in from_claus varchar(256) default null
	)
begin
	declare sqls varchar(3000);

	set sqls = 'delete ' + table_name;
	
	if from_claus is not null and char_length(from_claus) > 0 then 
		set sqls = sqls 
			+ ' from ' + from_claus;
	end if;
	set sqls = sqls 
   		+ ' where ' + where_claus;


   	execute immediate sqls;
end;



if exists (select 1 from sysprocedure where proc_name = 'slave_count_delete') then
	drop procedure slave_count_delete;
end if;

create 
	procedure slave_count_delete(
		out deleted integer
		, in table_name varchar(50)
		, in where_claus varchar(1000)
		, in from_claus varchar(256) default null
	)
begin
	declare sqls varchar(3000);

	set sqls = 'delete ' + table_name;
	
	if from_claus is not null and char_length(from_claus) > 0 then 
		set sqls = sqls 
			+ ' from ' + from_claus;
	end if;
	set sqls = sqls 
   		+ ' where ' + where_claus;

   	execute immediate sqls;
   	select @@rowcount into deleted;
end;




--****************************************************************
--                               INSERT
--****************************************************************
if exists (select 1 from sysprocedure where proc_name = 'slave_insert') then
	drop function slave_insert;
end if;

create 
	procedure slave_insert(
		  in table_name varchar(50)
		, in field_claus varchar(256) default null
		, in values_claus varchar(1000) default null
		, in select_claus varchar(1000) default null
	)
begin
	
	declare sqls varchar(3000);
	set sqls = 'insert into ' + table_name;
	if char_length(field_claus) > 0 then 
		set sqls = sqls + ' (' + field_claus + ')';
	end if;

	if char_length(values_claus) > 0 then 
		set sqls = sqls 
			+ ' values (' + values_claus + ')';
	elseif char_length(select_claus) > 0 then
		set sqls = sqls 
			+ ' ' + select_claus;
	end if;

	--raiserror 17000 'sqls = "%1!"', sqls
   	execute immediate  sqls;
end;



if exists (select 1 from sysprocedure where proc_name = 'slave_count_insert') then
	drop procedure slave_count_insert;
end if;

create 
	procedure slave_count_insert(
		  out inserted integer
		, in table_name varchar(50)
		, in field_claus varchar(256) default null
		, in values_claus varchar(1000) default null
		, in select_claus varchar(1000) default null
	)
begin
	
	declare sqls varchar(3000);
	declare inserted integer;
	declare f_return_id integer;
	set f_return_id = 0;

	set sqls = 'insert into ' + table_name;
	if char_length(field_claus) > 0 then 
		set sqls = sqls + ' (' + field_claus + ')';
	end if;

	if char_length(values_claus) > 0 then 
		set sqls = sqls 
			+ ' values (' + values_claus + ')';
		set f_return_id = 1;
	elseif char_length(select_claus) > 0 then
		set sqls = sqls 
			+ ' ' + select_claus;
	end if;

	--raiserror 17000 'sqls = "%1!"', sqls
	execute immediate  sqls;
	select @@rowcount into inserted;
	if f_return_id = 1 then
		set inserted = @id;
	end if;
end;

--****************************************************************
--                               UPDATE
--****************************************************************
if exists (select 1 from sysprocedure where proc_name = 'slave_update') then
	drop function slave_update;
end if;

create 
	procedure slave_update(
		  in table_name varchar(50)
		, in field_claus varchar(256)
		, in value_claus varchar(1000)
		, in where_claus varchar(1000)
		, in from_claus varchar(256) default null
	)
begin
	
	declare sqls varchar(3000);
	
   	set sqls =  'update ' + table_name + ' set ' + field_claus
   		+ ' = ' + value_claus;
	if from_claus is not null and char_length(from_claus) > 0 then 
		set sqls = sqls 
			+ ' from ' + from_claus;
	end if;
	set sqls = sqls 
   		+ ' where ' + where_claus;

   	execute immediate sqls;
end;


if exists (select 1 from sysprocedure where proc_name = 'slave_count_update') then
	drop procedure slave_count_update;
end if;

create 
	procedure slave_count_update(
		  out updated integer
		, in table_name varchar(50)
		, in field_claus varchar(256)
		, in value_claus varchar(1000)
		, in where_claus varchar(1000)
		, in from_claus varchar(256) default null
	)
begin
	declare sqls varchar(3000);
	
   	set sqls =  'update ' + table_name + ' set ' + field_claus
   		+ ' = ' + value_claus;
	if from_claus is not null and char_length(from_claus) > 0 then 
		set sqls = sqls 
			+ ' from ' + from_claus;
	end if;
	set sqls = sqls 
   		+ ' where ' + where_claus;
   	execute immediate sqls ;
   	select @@rowcount into updated;
end;



--****************************************************************
--                               SELECT
--****************************************************************
/*
if exists (select 1 from sysprocedure where proc_name = 'slave_select') then
	drop function slave_select;
end if;

create function slave_select (in table_name char(50), in field_claus char(256), in where_claus varchar(1000)) 
returns varchar(1000)
begin
	declare ret varchar(1000);
	call slave_select_prior (ret, table_name, field_claus, where_claus);
	return ret;
end;
*/


if exists (select '*' from sysprocedure where proc_name like 'slave_select') then  
	drop procedure slave_select;
end if;

create 
	procedure slave_select (
		  out selected varchar(1000)
		, in table_name varchar(50)
		, in field_claus varchar(256)
		, in where_claus varchar(1000) default null
		, in join_claus varchar(1000) default null
	)
begin
--	declare ret varchar(1000);
--	declare mValue varchar(1000);
    declare sqls varchar(2000);
    set sqls = 
   	  'select ' + field_claus 
   		+ ' into selected '
   		+ ' from '  + table_name ;

   	if join_claus is not null then
	    set sqls = sqls
   			+' '+ join_claus;
   	end if;

   	if where_claus is not null then
	    set sqls = sqls
   			+ ' where ' + where_claus;
   	end if;

   	execute immediate sqls;
end;


if exists (select '*' from sysprocedure where proc_name like 'slave_block_table') then  
	drop procedure slave_block_table;
end if;

create 
	procedure slave_block_table (
		  in emitter_server_name varchar(1000)
		, in block_table_name varchar(50)
	)
begin
	declare var_name varchar(100);
	declare var_value char(1);

	set var_name = make_block_name (emitter_server_name, block_table_name);

//	execute immediate 'select ' + var_name + ' into var_value' ;
	if var_value = 1 then
		raiserror 17010 'Переменная %1! уже была определена. \n Запомните условия возникновения этой ошибки и сообщите администратору', var_name;
	end if;

	execute immediate 'create variable ' + var_name +' integer';
	execute immediate 'set ' + var_name + ' = 1';

end;


if exists (select '*' from sysprocedure where proc_name like 'slave_unblock_table') then  
	drop procedure slave_unblock_table;
end if;

create 
	procedure slave_unblock_table (
		  in emitter_server_name varchar(50)
		, in unblock_table_name varchar(50)
	)
begin
	declare var_name varchar(100);
	declare var_value char(1);

	set var_name = make_block_name (emitter_server_name, unblock_table_name);
//	execute immediate 'select ' + var_name + ' into var_value' ;
	if var_value = 1 then
		raiserror 17010 'Ошибка при разблокировке таблицы.\n Переменная %1! НЕ была определена. \n Запомните условия возникновения этой ошибки и сообщите администратору', var_name;
	end if;

	execute immediate 'drop variable ' + var_name;

end;


if exists (select '*' from sysprocedure where proc_name like 'make_block_name') then  
	drop function make_block_name;
end if;

create 
	function make_block_name (
		  in emitter_server_name varchar(50)
		, in table_name varchar(50)
	)
	returns varchar(100)
begin
	return '@'+emitter_server_name + '_' + table_name;

end;



