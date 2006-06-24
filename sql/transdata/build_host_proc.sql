--****************************************************************
--                      BUILD HOST PROCEDURE
--****************************************************************

if exists (select '*' from sysprocedure where proc_name like 'build_remote_host') then
	drop procedure build_remote_host;
end if;

create procedure build_remote_host (
	  p_proc_name varchar(100)
	, p_params varchar(2000)
)
begin
	declare v_proc_name varchar(64);
	declare v_rp_proc_name varchar(64);
	declare v_rhost_proc_name varchar(64);
	declare v_strip_params varchar(1000);
	declare v_sql varchar(8000);
	declare v_params_len integer;

--	declare v_proc_name_remote varchar(64);

	set v_proc_name = p_proc_name + '_remote';

	for c_rmt_servers as a dynamic scroll cursor for
		select srvname as r_slave_server from sys.sysservers 
	do

		call build_rp_procedure (
			  r_slave_server
			, p_proc_name
			, p_params
		);

	end for;


	set v_params_len = char_length(trim(p_params));
	set v_strip_params = parse_params(p_params);

	if exists (select 1 from sysprocedure where proc_name like v_proc_name ) then 
		execute immediate 
			'drop procedure ' + v_proc_name;
	end if;
	
	set v_sql =
		  '\ncreate PROCEDURE ' + v_proc_name
		+ '\n('
		+ '\n	p_sysname varchar(64)'
	;
	if v_params_len > 0 then
		set v_sql = v_sql
			+', ' + 	p_params;
	end if;
	set v_sql = v_sql +
		+ '\n)'
		+ '\nbegin'
		+ '\n	declare v_sysname varchar(64);'
		+ '\n	declare v_proc_name varchar(64);'
		+ '\n	select srvname into v_sysname from sys.sysservers where srvname = p_sysname;'
		+ '\n	--message ''v_sysname = ''+ v_sysname to client;'
		+ '\n	if v_sysname is null then raiserror 17000, ''Error in remote procedure %1!'', ' + v_proc_name + '; end if;'
		+ '\n	set v_proc_name = ''' + p_proc_name + '_'' + p_sysname;'
		+ '\n	--message ''v_proc_name = ''+ v_proc_name to client;'
		+ '\n	execute immediate ''call '' + v_proc_name '
	;

	if v_params_len > 0 then
		set v_sql = v_sql
			+' + ''( ' + v_strip_params + ')''';
	end if;
	set v_sql = v_sql +
		+ '\n'
		+ '\nend;'
	;

	message v_sql to client;
	
	
	execute immediate v_sql;

end;




if exists (select '*' from sysprocedure where proc_name like 'build_rp_procedure') then
	drop procedure build_rp_procedure;
end if;

create procedure build_rp_procedure(
	  p_slave_server varchar(30)
	, p_proc_name varchar(100)
	, p_params varchar(2000)
	, p_checking integer default 1
)
begin
	declare v_proc_name varchar(64);

	declare v_rp_proc_name varchar(64);

	declare v_strip_params varchar(1000);
	declare v_sql varchar(3000);

	set v_proc_name = p_proc_name +'_' + p_slave_server;
	set v_rp_proc_name = 'r_' + v_proc_name;

	if exists (select 1 from sysprocedure where proc_name like v_rp_proc_name ) then 
		execute immediate 
			'drop procedure ' + v_rp_proc_name;
    end if;
	
	execute immediate 
		  ' create procedure ' + v_rp_proc_name
		+ ' ('
		+ 	p_params
		+ ' )'
		+ ' at '''+p_slave_server+';;;' + p_proc_name + '''';


	if exists (select 1 from sysprocedure where proc_name like v_proc_name ) then 
		execute immediate 
			'drop procedure ' + v_proc_name;
	end if;

	set v_strip_params = parse_params(p_params);

	set v_sql =
		  '\ncreate PROCEDURE ' + v_proc_name
		+ '\n('
		+ 	p_params
		+ '\n)'
		+ '\nbegin'
	;
	if p_checking = 1 then
		set v_sql = v_sql
			+ '\n	if get_standalone('''+ p_slave_server + ''') = 0 then'
		;
	end if;

	set v_sql = v_sql
		+ '\n		call ' + v_rp_proc_name + '( ' +  v_strip_params + ');'
	;

	if p_checking = 1 then
		set v_sql = v_sql
			+ '\n	else'
			+ '\n		--call log_debug (''function '+ v_proc_name +''');'
			+ '\n	end if;'
		;
	end if;

	set v_sql = v_sql
		+ '\nend'
	;

	execute immediate v_sql;
end;





if exists (select '*' from sysprocedure where proc_name like 'build_host_procedure') then
	drop procedure build_host_procedure;
end if;

create procedure build_host_procedure(
	p_proc_name varchar(100)
	, p_params varchar(2000)
	, p_checking integer default 1
)
begin
	declare v_slave_proc_name varchar(64);

	set v_slave_proc_name = 'slave_' + p_proc_name;
	for c_rmt_servers as a dynamic scroll cursor for
		select srvname as r_slave_server from sys.sysservers 
	do

		call build_rp_procedure (
			  r_slave_server
			, v_slave_proc_name
			, p_params
			, p_checking
		);

	end for;
end;




if exists (select '*' from sysprocedure where proc_name like 'parse_params') then
	drop function parse_params;
end if;

create function parse_params(
	p_params varchar(1000)
) returns varchar(1000)
begin
	declare unikey varchar(1000);
	declare v_nomnom varchar(50);
	declare v_ret varchar(1000);
	declare comma char(1);
	declare p integer;

	set unikey = p_params;
	set parse_params = '';
	set comma = '';
	recurse: loop
		set p=charindex(',',unikey);
		if p = 0 then
			set v_nomnom=unikey;
			set unikey='';
		else
			set v_nomnom=substring(unikey,1,p-1);
			set unikey=substring(unikey,p+1);
		end if; 

		set parse_params = parse_params  
			+ comma + strip_param(v_nomnom);

		if char_length(unikey) <= 0 then
		  leave recurse;
		end if;
		set comma = ', ';
	end loop recurse;

	
end;


if exists (select '*' from sysprocedure where proc_name like 'strip_param') then
	drop function strip_param;
end if;

create function strip_param(
	p_param varchar(100)
) returns varchar(100)
begin

	declare var_ref1 varchar(8);
	declare var_ref2 varchar(8);
	declare var_ref3 varchar(8);
	declare space1   varchar(8);
	declare space2   varchar(8);
	declare space3   varchar(8);
	declare space4   varchar(8);
	declare vlen integer;
	declare v_found bit;
	declare v_first bit;
	declare v_first_word varchar(100);

	set strip_param = p_param;
	set var_ref1 = 'inout';
	set var_ref2 = 'in';
	set var_ref3 = 'out';
	set space1 = ' ';
	set space2 = char(9);
	set space3 = char(10);
	set space4 = char(13);

	set strip_param = replace(strip_param, space2, space1);
	set strip_param = replace(strip_param, space3, space1);
	set strip_param = replace(strip_param, space4, space1);
	set strip_param = ltrim(strip_param);

//	set v_first_word=substr(strip_param, 1, charindex(space1,strip_param));

	

	set v_found = 0;
	set vlen = start_with(strip_param, var_ref1); 
	if (v_found = 0 and vlen != 0) then 
		set v_first_word = substring(strip_param, vlen + 1, 1);
		message 'v_first_word = ''', v_first_word, ''''  to client;
		if (v_first_word = space1) then
			set strip_param = substring(strip_param, vlen + 1); 
			set v_found = 1; 
		end if;
	end if;

	set vlen = start_with(strip_param, var_ref2); 
	if (v_found = 0 and vlen != 0) then 
		set v_first_word = substring(strip_param, vlen + 1, 1);
		message 'v_first_word = ''', v_first_word, ''''  to client;
		if (v_first_word = space1) then
			set strip_param = substring(strip_param, vlen + 1); 
			set v_found = 1; 
		end if;
	end if;

	set vlen = start_with(strip_param, var_ref3); 
	if (v_found = 0 and vlen != 0) then 
		set v_first_word = substring(strip_param, vlen + 1, 1);
		message 'v_first_word = ''', v_first_word, ''''  to client;
		if (v_first_word = space1) then
			set strip_param = substring(strip_param, vlen + 1); 
			set v_found = 1; 
		end if;
	end if;

	set strip_param = ltrim(strip_param);

	set strip_param = substring(strip_param, 1, charindex(space1, strip_param)-1);


end;


if exists (select '*' from sysprocedure where proc_name like 'start_with') then
	drop function start_with;
end if;

create function start_with(
	p_token varchar(100)
	, p_pattern varchar(50)
) returns integer
begin
	if p_token = '' or p_token is null then
		return 0;
	end if;
	set start_with = char_length(p_pattern);
//	message 'start_with = ', start_with to client;
//	message 'left(p_token, start_with) = ', left(p_token, start_with) to client;
	if left(p_token, start_with) != p_pattern then
	    set start_with = 0;
	end if;
end;


/*
if exists (select '*' from sysprocedure where proc_name like 'build_remote_procedure') then
	drop procedure build_remote_procedure;
end if;

create procedure build_remote_procedure(
		p_proc_name varchar(100)
		, p_params varchar(2000)
	)
begin
	declare v_proc_name varchar(64);
	declare v_slave_name varchar(64);
//	declare p_proc_name varchar(64);
//	declare v_slave_server varchar(32);
	declare v_remote_proc_name varchar(64);

//	set p_proc_name = 'currency_rate';

//		set v_slave_server = 'mm';
		set v_remote_proc_name = p_proc_name + '_remote';
--		set v_proc_name = v_remote_proc_name +'_' + v_slave_server;
    
		if exists (select 1 from sysprocedure where proc_name like v_proc_name ) then 
			execute immediate 
				'drop procedure ' + v_proc_name;
        end if;
		
		execute immediate 
			  ' create PROCEDURE ' + v_proc_name
			+ ' ( p_server, '
			+ p_params
			+ ' )'
			+ '\n begin'
			+ '\n    execute immedate ''call '+v_remote_proc_name + '_ + p_server'

			+ '\n end;'
		:
end;

*/


if exists (select '*' from sysprocedure where proc_name like 'build_remote_table') then
	drop procedure build_remote_table;
end if;

create procedure build_remote_table(
		p_table_name varchar(100)
		, f_create integer default 1
	)
begin
	declare v_table_name varchar(64);

	for c_rmt_servers as a dynamic scroll cursor for
		select srvname as r_slave_server from sys.sysservers 
	do

		set v_table_name = p_table_name +'_' + r_slave_server;
    
        if f_create = 0 then
			if exists (select 1 from systable where table_name = v_table_name ) then 
				execute immediate 
					'drop table ' + v_table_name;
            end if;
        else 
			if not exists (select 1 from systable where table_name = v_table_name ) then 
				execute immediate 
					'create existing table ' + v_table_name + ' at '''+ r_slave_server +'...'+ p_table_name + '''';
            end if;
		end if;
		
	end for;
end;


