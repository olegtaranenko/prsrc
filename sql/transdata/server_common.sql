--****************************************************************
--                               INSERT
--****************************************************************
if exists (select '*' from sysprocedure where proc_name like 'insert_remote') then  
	drop procedure insert_remote;
end if;

create 
	procedure insert_remote(
		  in server_name varchar(32)
		, in table_name char(50)
		, in field_claus char(256) default null
		, in values_claus char(1000) default null 
		, in select_claus char(1000) default null
	)
begin
	
	declare sqls varchar(3000);
	set sqls = '('''  + table_name + '''';
	if field_claus is not null then
		set sqls = sqls + ', ''' + field_claus + '''';
	else 
		set sqls = sqls + ', null';
	end if;
	
	if values_claus is not null then
		set sqls = sqls + ', ''' + values_claus + '''' ;
	else 
		set sqls = sqls + ', null';
	end if;

	if select_claus is not null then
		set sqls = sqls + ', ''' + select_claus + '''' ;
	else 
		set sqls = sqls + ', null';
	end if;
	
	set sqls = sqls + ')';
	// call remote procedure
	message sqls to client;
	execute immediate 'call slave_insert_' + server_name + sqls;
end;




if exists (select '*' from sysprocedure where proc_name like 'insert_count_remote') then  
	drop function insert_count_remote;
end if;

create 
	function insert_count_remote(
			in server_name varchar(32)
			, in table_name char(50)
			, in field_claus char(256) default null
			, in values_claus char(1000) default null 
			, in select_claus char(1000) default null
	)
		returns integer
begin
	
	declare sqls varchar(3000);
	declare inserted integer;

	set sqls = '(inserted, '''  + table_name + '''';
	if field_claus is not null then
		set sqls = sqls + ', ''' + field_claus + '''';
	else 
		set sqls = sqls + ', null';
	end if;
	
	if values_claus is not null then
		set sqls = sqls + ', ''' + values_claus + '''' ;
	else 
		set sqls = sqls + ', null';
	end if;

	if select_claus is not null then
		set sqls = sqls + ', ''' + select_claus + '''' ;
	else 
		set sqls = sqls + ', null';
	end if;
	
	
	set sqls = sqls + ')';
	
   	execute immediate 'call slave_count_insert_' + server_name + sqls;
   	return inserted;
end;



--****************************************************************
--                               UPDATE
--****************************************************************

if exists (select '*' from sysprocedure where proc_name like 'update_remote') then 
	drop procedure update_remote;
end if;
create 
	procedure update_remote(
		in server_name varchar(32)
		, in table_name varchar(50)
		, in field_claus varchar(256)
		, in value_claus varchar(1000) 
		, in where_claus varchar(1000)
		, in from_claus varchar(256) default null
	)
begin
	
	declare sqls varchar(3000);
	set sqls = '('''  + table_name + '''';
	if field_claus is not null then
		set sqls = sqls + ', ''' + field_claus + '''';
	else 
		set sqls = sqls + ', null';
	end if;
	
	if value_claus is not null then
		set sqls = sqls + ', ''' + value_claus + '''' ;
	else 
		set sqls = sqls + ', null';
	end if;

	if where_claus is not null then
		set sqls = sqls + ', ''' + where_claus + '''' ;
	else 
		set sqls = sqls + ', null';
	end if;
	
	if from_claus is not null then
		set sqls = sqls + ', ''' + from_claus + '''' ;
	else 
		set sqls = sqls + ', null';
	end if;
	
	set sqls = sqls + ')';
	
	message sqls to client;

   	execute immediate 'call slave_update_' + server_name + sqls;
end;



if exists (select '*' from sysprocedure where proc_name like 'update_count_remote') then 
	drop function update_count_remote;
end if;

create 
	function update_count_remote(
		in server_name varchar(32)
		, in table_name varchar(50)
		, in field_claus varchar(256)
		, in value_claus varchar(1000) 
		, in where_claus varchar(1000)
		, in join_claus varchar(256) default null
	)
		returns integer
begin
	
	declare sqls varchar(3000);
	declare updated integer;

	set sqls = '(updated, '''  + table_name + '''';
	if field_claus is not null then
		set sqls = sqls + ', ''' + field_claus + '''';
	else 
		set sqls = sqls + ', null';
	end if;
	
	if value_claus is not null then
		set sqls = sqls + ', ''' + value_claus + '''' ;
	else 
		set sqls = sqls + ', null';
	end if;

	if where_claus is not null then
		set sqls = sqls + ', ''' + where_claus + '''' ;
	else 
		set sqls = sqls + ', null';
	end if;

	if join_claus is not null then
		set sqls = sqls + ', ''' + join_claus + '''' ;
	else 
		set sqls = sqls + ', null';
	end if;
	
	set sqls = sqls + ')';
	message 'update_count_remote: ', sqls to client;
	
   	execute immediate 'call slave_count_update_' + server_name + sqls;
	message 'updated_count_remote: updated = ', updated to client;
   	return updated;

end;




--****************************************************************
--                               DELETE
--****************************************************************

if exists (select '*' from sysprocedure where proc_name like 'delete_remote') then  
	drop procedure delete_remote;
end if;


create 
	procedure delete_remote(
			in server_name varchar(32)
			, in table_name varchar(50)
			, in where_cond varchar(1000)
			, in join_cond varchar(1000) default null
	)
begin
	declare sqls varchar(2000);
	set sqls = 
   		'call slave_delete_' + server_name 
   			+ '('''  
   				+ table_name 
   				+ ''', ''' + where_cond + ''''
   	;
   	if join_cond is not null and char_length(join_cond) > 0 then
   		set sqls = sqls 
   			+ ', ''' + join_cond + '''';
   	else 
   		set sqls = sqls 
   			+ ', null';
   	end if;

	set sqls = sqls 
		+ ')';

   	execute immediate sqls;

end;



----------------
if exists (select '*' from sysprocedure where proc_name like 'delete_count_remote') then  
	drop function delete_count_remote;
end if;

create 
	function delete_count_remote(
			in server_name varchar(32)
			, in table_name varchar(50)
			, in where_cond varchar(1000)
	)
		returns integer
begin
	declare sqls varchar(2000);
	declare deleted integer;

	set sqls = 
		'call slave_count_delete_' 
		+ server_name 
		+ '(deleted, ''' + table_name 
		+ ''', ''' + where_cond 
		+ ''')'
	;

   	execute immediate sqls;

   	return deleted;
 	
end;




--****************************************************************
--                               SELECT
--****************************************************************


if exists (select '*' from sysprocedure where proc_name like 'select_remote') then  
	drop function select_remote;
end if;

create 
	function select_remote(
		in server_name varchar(32)
		, in table_name varchar(50)
		, in field_claus varchar(256)
		, in where_claus varchar(1000) default null
		, in join_claus varchar(1000) default null
	)
	returns varchar(2000)
begin
	declare selected varchar(2000);
	declare q varchar(2000);

	set q = 
		'call slave_select_' + server_name 
			+ '('
			+ ' selected '
			+ ', ' + '''' + table_name  + ''''
			+ ', ' + '''' + field_claus + ''''
	;

	if where_claus is not null then
		set q = q 
			+ ', ' + '''' + where_claus + ''''
	end if; 

	if join_claus is not null then
		set q = q 
			+ ', ' + '''' + join_claus + ''''
	end if; 

	set q = q +')';

	message q to client;
	execute immediate q;

	return selected;
end;

--****************************************************************
--                               CALL
--****************************************************************


if exists (select '*' from sysprocedure where proc_name like 'call_remote') then
	drop function call_remote;
end if;

create procedure call_remote(
		  p_server varchar(50)
		, p_proc_name varchar(100)
		, p_params varchar(2000)
	)

begin
	declare v_sql varchar(254);

	set v_sql = 'call ' + p_proc_name + '_'+ p_server;
	if p_params is not null then
		set v_sql = v_sql + '(' + p_params + ')';
	else 
		set v_sql = v_sql + '( )';
	end if;
//	message v_sql to client;
	execute immediate v_sql;
end;


--****************************************************************
--                        BLOCK REMOTE
--****************************************************************


if exists (select '*' from sysprocedure where proc_name like 'block_remote') then
	drop function block_remote;
end if;

create procedure block_remote(
		  p_server varchar(50)
		, p_caller varchar(50)
		, p_table_name varchar(100)
	)

begin
	declare v_sql varchar(254);
	declare sync char(1);
	declare first_time integer;

	declare blocks_inited integer;

	set first_time = 0;

	rep:
	--loop
		begin
			execute immediate 'set blocks_inited = @blocks_inited';
			--leave rep;
		exception when others then
			message 'Server ', get_server_name(), ' :: Launch bootstrap_blocking()' to log;
			if first_time = 0 then
				set first_time = 1;
				call bootstrap_blocking();
				waitfor delay convert(time, '00:00:00.500');
				message '  ...  bootstrap_blocking ended successfully' to log;
			end if;
			--leave rep;
		end;
	--end loop;

	set v_sql = 'call slave_block_table_'+ p_server + '(sync, '''+ p_caller+ ''', '''+p_table_name+''')';
	message v_sql to log;
	execute immediate v_sql;
	--waitfor delay convert(time, '00:00:00.500');
end;


if exists (select '*' from sysprocedure where proc_name like 'unblock_remote') then
	drop function unblock_remote;
end if;

create procedure unblock_remote(
		  p_server varchar(50)
		, p_caller varchar(50)
		, p_table_name varchar(100)
	)

begin
	declare v_sql varchar(254);
	declare sync char(1);

	begin
		set v_sql = 'call slave_unblock_table_'+ p_server + '(sync, '''+ p_caller+ ''', '''+p_table_name+''')';
		execute immediate v_sql;
		exception when others then
			-- do nothing
	end;
end;





call build_host_procedure (
		'unblock_table'
		,	'out sync integer'
		+ ', in emitter_server_name char(50)'
		+ ', in block_table_name char(50)'
);



call build_host_procedure ( 
			'block_table'
		,	'out sync integer'
		+ ', in emitter_server_name char(50)'
		+ ', in block_table_name char(50)'
);



call build_host_procedure ( 
		 'select'
	 	, '  out selected char(1000)'
		+ ', in table_name char(50)'
		+ ', in field_claus char(256)'
		+ ', in where_claus char(1000) default null'
		+ ', in join_claus char(1000) default null'
);

call build_host_procedure ( 
		'count_update'
		, 'out updated integer'
		+ ', in table_name char(50)'
		+ ', in field_claus char(256) '
		+ ', in values_claus char(1000) '
		+ ', in where_claus char(1000)'
		+ ', in join_claus char(256) default null'
);


call build_host_procedure ( 
		'update'
		, 'in table_name char(50)'
		+ ', in field_claus char(256) '
		+ ', in values_claus char(1000) '
		+ ', in where_claus char(1000)'
		+ ', in join_claus char(256) default null'
);	



call build_host_procedure ( 
		'count_insert'
		, '  out inserted integer'
		+ ', in table_name char(50)'
		+ ', in field_claus char(256) default null'
		+ ', in values_claus char(2000) default null '
		+ ', in select_claus char(1000) default null'
);


call build_host_procedure ( 
		 'insert'
		, '	in table_name char(50)'
		+ '	, in field_claus char(256) default null'
		+ '	, in values_claus char(1000) default null '
		+ '	, in select_claus char(1000) default null'
);

call build_host_procedure ( 
		 'count_delete'
		, '	out deleted integer'
		+ '	,in table_name char(50)'
		+ '	,in where_cond char(2000)'
		+ ', in join_claus char(256) default null'
);


call build_host_procedure ( 
		 'delete'
		, 'in table_name char(50)'
		+ ',in where_cond char(2000)'
		+ ', in join_claus char(256) default null'
);

call build_host_procedure (
			'cre_block_var'
		, '  in var_name char(100)'
);

