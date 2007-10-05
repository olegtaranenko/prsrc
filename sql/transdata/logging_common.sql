--****************************************************************
--                               LOGGING
--****************************************************************


if exists (select '*' from sysprocedure where proc_name like 'show_monitor') then  
	drop procedure show_monitor;
end if;

create 
	procedure show_monitor(msg varchar(1000), monitor varchar(20))
begin

	if monitor = 'all' or monitor = 'client' then
		message msg to client;
	end if;
	if monitor = 'all' or monitor = 'console' then
		message msg to console;
	else 
		message msg to log;
	end if;
end;

if exists (select '*' from sysprocedure where proc_name like 'log_error') then  
	drop procedure log_error;
end if;

create 
	procedure log_error(msg varchar(1000), monitor varchar(20) default 'log')
begin
	call show_monitor ('ERROR [' + current database + '@' + get_server_name() + ']: ' + msg, monitor);
	if convert(integer, sqlcode) != 0 then
		call show_monitor('	SQLCODE =' + convert(varchar(20), SQLCODE) + ', SQLSTATE = ' + convert(varchar(20),SQLSTATE)
			,monitor);
	end if;
end;



if exists (select '*' from sysprocedure where proc_name like 'log_critical') then  
	drop procedure log_critical;
end if;

create 
	procedure log_critical(msg varchar(1000), monitor varchar(20) default 'all')
begin
	call show_monitor ('CRITICAL [' + current database + '@' + get_server_name() + ']: ' + msg, monitor);
	if convert(integer, sqlcode) != 0 then
		call show_monitor(
			'	SQLCODE =' + convert(varchar(20), SQLCODE) 
			+ ', SQLSTATE = ' + convert(varchar(20),SQLSTATE)
			,monitor);
	end if;
end;

if exists (select '*' from sysprocedure where proc_name like 'log_warning') then  
	drop procedure log_warning;
end if;

create 
	procedure log_warning(msg varchar(1000), monitor varchar(20) default 'log')
begin
	call show_monitor ('WARNING [' + current database + '@' + get_server_name() + ']: ' + msg, monitor);
end;

if exists (select '*' from sysprocedure where proc_name like 'log_debug') then  
	drop procedure log_debug;
end if;

create 
	procedure log_debug(msg varchar(1000), monitor varchar(20) default 'log')
begin
	if (@debug = 1) then
		call show_monitor ('DEBUG [' + current database + '@' + get_server_name() + ']: ' + msg, monitor);
	end if;
	exception when others then
		return;
end;



