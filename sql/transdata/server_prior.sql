--****************************************************************
--                               INSERT
--****************************************************************

if exists (select 1 from sysprocedure where proc_name = 'insert_host') then
	drop function insert_host;
end if;

create 
	procedure insert_host(
		in table_name char(50)
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
	
	  for v_remote_name as a dynamic scroll cursor for
		select srvname as cur_remote from sys.sysservers s join guideventure v on s.srvname = v.sysname and v.standalone = 0 do
		
    	execute immediate 'call slave_insert_' + cur_remote + sqls;
	  end for;
end;




if exists (select 1 from sysprocedure where proc_name = 'insert_count_host') then
	drop function insert_count_host;
end if;

create 
	function insert_count_host(
			in table_name char(50)
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
	
	  for v_remote_name as a dynamic scroll cursor for
		select srvname as cur_remote from sys.sysservers s join guideventure v on s.srvname = v.sysname and v.standalone = 0 do
		
    	execute immediate 'call slave_count_insert' + cur_remote + sqls;
	  end for;

	return inserted;
end;



--****************************************************************
--                               UPDATE
--****************************************************************

if exists (select '*' from sysprocedure where proc_name like 'update_host') then 
	drop procedure update_host;
end if;
create 
	procedure update_host(
		in table_name varchar(50)
		, in field_claus varchar(256)
		, in value_claus varchar(1000) 
		, in where_claus varchar(1000)
		, in join_claus varchar(256) default null
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
	
	if join_claus is not null then
		set sqls = sqls + ', ''' + join_claus + '''' ;
	else 
		set sqls = sqls + ', null';
	end if;
	
	set sqls = sqls + ')';
	
	  for v_remote_name as a dynamic scroll cursor for
		select srvname as cur_remote from sys.sysservers s join guideventure v on s.srvname = v.sysname and v.standalone = 0 do
		
    	execute immediate 'call slave_update_' + cur_remote + sqls;
	  end for;
end;


if exists (select '*' from sysprocedure where proc_name like 'update_count_host') then 
	drop function update_count_host;
end if;
create 
	function update_count_host(
		in table_name varchar(50)
		, in field_claus varchar(256)
		, in value_claus varchar(1000)
		, in where_claus varchar(1000)
		, in join_claus varchar(256) default null
	) 
		returns integer
begin
	
	declare sqls varchar(3000);
	declare updated integer;
	declare permit bit;

	set sqls = '(';  
	set sqls = sqls + 'updated';
	set sqls = sqls + ', ''' + table_name + '''';
	

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
	
	set updated = -1;

	for v_remote_name as rnm dynamic scroll cursor for
		select srvname as cur_remote from sys.sysservers s join guideventure v on s.srvname = v.sysname and v.standalone = 0 
	do

//		if wf_permit_session(table_name, 'update') = 1 then
    		execute immediate 'call slave_count_update_' + cur_remote + sqls;
//    	else 
//    		call log_error(' server ' + v_remote_server + ' is configured as not accessible');
//    		set updated = 0;
//    	end if;
	end for;
	return updated;
end;



--****************************************************************
--                               DELETE
--****************************************************************

if exists (select 1 from sysprocedure where proc_name = 'delete_host') then
	drop function delete_host;
end if;

create 
	procedure delete_host(in table_name varchar(50), in where_cond varchar(1000))
begin
	
	  for v_remote_name as a dynamic scroll cursor for
		select srvname as cur_remote from sys.sysservers s join guideventure v on s.srvname = v.sysname and v.standalone = 0 do
--		raiserror 17002 'call slave_delete_' + cur_remote + '('''  + table_name + ''', ''' + where_cond + ''')';
		
    	execute immediate 'call slave_delete_' + cur_remote + '('''  + table_name + ''', ''' + where_cond + ''')';
	  end for;
end;



----------------
if exists (select 1 from sysprocedure where proc_name = 'delete_count_host') then
	drop function delete_count_host;
end if;

create 
	function delete_count_host(
			in table_name varchar(50)
			, in where_cond varchar(1000)
	)
		returns integer
begin
	
	declare sqls varchar(2000);
	declare deleted integer;


	for v_remote_name as a dynamic scroll cursor for
		select srvname as cur_remote from sys.sysservers s join guideventure v on s.srvname = v.sysname and v.standalone = 0 do


		set sqls = 
			'call slave_delete_' 
			+ cur_remote
			+ '(deleted'
			+ ', ''' + table_name  + ''''
			+ ', ''' + where_cond + ''''
			+ ')'
		;

		execute immediate sqls;
	end for;
	return deleted;
end;




if exists (select '*' from sysprocedure where proc_name like 'call_host') then
	drop function call_host;
end if;

create procedure call_host(
		p_proc_name varchar(100)
		, p_params varchar(2000)
	)

begin
	declare v_sql varchar(254);

	for v_remote_name as a dynamic scroll cursor for
		select srvname as cur_remote from sys.sysservers s join guideventure v on s.srvname = v.sysname and v.standalone = 0 
	do
		set v_sql = 'call slave_' + p_proc_name + '_'+ cur_remote + '(' + p_params + ')';
		message v_sql to client;
		execute immediate v_sql;
	end for;
end;



call build_host_procedure (
		 'set_standalone'
		, '  out p_success integer'
		+ ', p_status char(23)'
		, 0 // generate no checkable
);

call build_host_procedure (
		 'get_standalone'
		, 'out p_standalone integer'
		, 0 // generate no checkable
);

