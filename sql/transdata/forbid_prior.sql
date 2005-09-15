if exists (select 1 from systable where table_name = 'zz_forbid_session') then
	drop table zz_forbid_session;
end if;

// 
create 
	GLOBAL TEMPORARY 
table zz_forbid_session (
	 server_name varchar(32) not null default ''
	, table_name varchar(64) not null default ''
	, operation varchar(32) not null default ''
	, fallback char(1)  not null default ''
	, table_field varchar(64) not null default ''
	, primary key (server_name, table_name, operation, fallback, table_field)
) 
on commit preserve rows
;



if exists (select 1 from systable where table_name = 'zz_forbid') then
	drop table zz_forbid;
end if;

// 
create 
	table zz_forbid (
	 server_name varchar(32) not null default ''
	, table_name varchar(64) not null default ''
	, operation varchar(32) not null default ''
	, fallback char(1)  not null default ''
	, table_field varchar(64) not null default ''
	, primary key (server_name, table_name, operation, fallback, table_field)
)
;

insert into zz_forbid (server_name, table_name, operation, fallback, table_field) 
select 'stime', 'mat', '', '', '' union
select 'stime', 'jmat', '', 'delete', 'y' union
select 'pm', '', '', '', '' union
select 'pm', 'mat', '', '', '';



--create unique index ui_forbid on zz_forbid(server_name, table_name, operation, fallback, table_field);


/*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/
/*                                                   */
/*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/
if exists (select 1 from sysprocedure where proc_name = 'wf_forbid') then
	drop function wf_forbid;
end if;

CREATE FUNCTION wf_forbid(
	p_server_name varchar(32)
	, p_table_name varchar(64) default ''
	, p_operation varchar(32) default ''
	, p_fallback char(1) default ''
	, p_table_field varchar(64) default ''
)
	returns bit
begin
	// по умолчанию, все, что не запрещено явно - разрешено
	set wf_forbid = 0;	

	if p_server_name is null then
		return;
	end if;


	for c_forbidden as cf dynamic scroll cursor for
		select 
			table_name as r_table_name
			, operation as r_operation
			, fallback as r_fallback
			, table_field as r_table_field
		from zz_forbid_session
		where 
				server_name = p_server_name
		order by 1, 2, 3, 4
	do
		-- с этим сервером вообще запрещены всякое общение
		if r_table_name = '' then
			return 1;
		end if;

		-- проверяем можно ли делать что-то именно с этой таблицей
		if p_table_name != '' then
			-- требуется проверка на конкретную таблицу
			-- если таблица не та, смотрим в следующие записи
			if p_table_name = r_table_name then
				-- проверяем уже на конкретную операцию по таблице
				if p_operation != '' then
					-- если опрации не та, смотрим дальше
					if p_operation = r_operation then
						return 1;
					end if;

					if p_table_field != '' then
						if p_table_field = r_table_field then
							return 1;
						end if;
			    	
					else
						return 0;
					end if;
				else 
					-- была затребовано проверка на конкретную 
					-- операцию по таблице
					-- возварщаем разрешение на операцию.
					-- дальше смотреть не нужно, потому что 
					-- результат отсортирован так, что null появляется первым
					return 0;
				end if;
			end if;
		else 
			// просто проверяем, можно ли работать с сервером
			return 0; // можно 
		end if;
	end for;
end;

if exists (select 1 from sysprocedure where proc_name = 'if_permit') then
	drop function if_permit;
end if;

CREATE FUNCTION if_permit(
	p_server_name varchar(32)
	, p_table_name varchar(64) default ''
	, p_operation varchar(32) default ''
	, p_fallback char(1) default ''
	, p_table_field varchar(64) default ''
)
	returns bit
begin
	if wf_forbid(p_server_name, p_table_name, p_operation, p_fallback, p_field_name) = 1 then 
		return 0 
	else 
		return 1 
	end if;
end;

/*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/
/*                                                   */
/*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/
if exists (select 1 from sysprocedure where proc_name = 'wf_load_session') then
	drop procedure wf_load_session;
end if;

CREATE procedure wf_load_session()
begin
	insert into zz_forbid_session (server_name, table_name, operation, fallback, table_field)
	on existing skip
	select server_name, table_name, operation, fallback, table_field from zz_forbid;
end;


/*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/
/*                                                   */
/*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/
if exists (select 1 from sysprocedure where proc_name = 'wf_create_forbid') then
	drop procedure wf_create_forbid;
end if;

CREATE procedure wf_create_forbid(
	  p_server_name varchar(32)
	, p_table_name varchar(64) default ''
	, p_operation varchar(32) default ''
	, p_fallback char(1) default '' // пока не используется
	, p_table_field varchar(64) default ''
	, p_persistent tinyint default 0 // 0 - temp only; 1 - persistent only; 2 - both
)
begin
	if p_persistent != 1 then
		insert into zz_forbid_session (server_name, table_name, operation, fallback, table_field)
		on existing skip
		select p_server_name, p_table_name, p_operation, p_fallback, p_field_name;
	elseif p_persistent > 0 then
		insert into zz_forbid (server_name, table_name, operation, fallback, table_field)
		on existing skip
		select p_server_name, p_table_name, p_operation, p_fallback, p_field_name;
	end if;
end;



if exists (select 1 from sysprocedure where proc_name = 'wf_release_forbid') then
	drop procedure wf_release_forbid;
end if;

CREATE procedure wf_release_forbid(
	  p_server_name varchar(32)
	, p_table_name varchar(64) default ''
	, p_operation varchar(32) default ''
	, p_fallback char(1) default '' // пока не используется
	, p_table_field varchar(64) default ''
	, p_persistent tinyint default 0 // 0 - temp only; 1 - persistent only; 2 - both
)
begin
	if p_persistent != 1 then
		delete from zz_forbid_session 
		where 
				isnull(p_server_name, '') = server_name
			and isnull(p_table_name, '' ) = table_name 
			and isnull(p_operation, ''  ) = operation
			and isnull(p_fallback, ''   ) = fallback;
	elseif p_persistent > 0 then
		delete from zz_forbid
		where 
				isnull(p_server_name, '') = server_name
			and isnull(p_table_name, '' ) = table_name 
			and isnull(p_operation, ''  ) = operation
			and isnull(p_fallback, ''   ) = fallback;
	end if;
end;



/*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/
/*                                                   */
/*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/
if exists (select 1 from sysprocedure where proc_name = 'wf_forbid_global') then
	drop procedure wf_forbid_global;
end if;

CREATE procedure wf_forbid_global(
	p_server_name varchar(32)
	, p_table_name varchar(64) default ''
	, p_operation varchar(32) default ''
	, p_table_field varchar(64) default ''
)
begin
	call wf_create_forbid (p_server_name, p_table_name, p_operation, null, p_field_name, 2);
end;


/*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/
/*                                                   */
/*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/
if exists (select 1 from sysprocedure where proc_name = 'wf_forbid_setting') then
	drop procedure wf_forbid_setting;
end if;

CREATE procedure wf_forbid_setting(
	p_server_name varchar(32)
	, p_table_name varchar(64) default ''
	, p_operation varchar(32) default ''
	, p_table_field varchar(64) default ''
)
begin
	call wf_create_forbid (p_server_name, p_table_name, p_operation, null, p_field_name, 1);
end;


/*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/
/*                                                   */
/*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/
if exists (select 1 from sysprocedure where proc_name = 'wf_forbid_session') then
	drop procedure wf_forbid_session;
end if;

CREATE procedure wf_forbid_session(
	p_server_name varchar(32)
	, p_table_name varchar(64) default ''
	, p_operation varchar(32) default ''
	, p_table_field varchar(64) default ''
)
begin
	call wf_create_forbid (p_server_name, p_table_name, p_operation, p_fallback, p_field_name, 0);
end;



/*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/
/*                                                   */
/*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/
if exists (select 1 from sysprocedure where proc_name = 'wf_permit_global') then
	drop procedure wf_permit_global;
end if;

CREATE procedure wf_permit_global(
	p_server_name varchar(32)
	, p_table_name varchar(64) default ''
	, p_operation varchar(32) default ''
	, p_table_field varchar(64) default ''
)
begin
	call wf_release_forbid (p_server_name, p_table_name, p_operation, p_fallback, p_field_name, 2);
end;


/*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/
/*                                                   */
/*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/
if exists (select 1 from sysprocedure where proc_name = 'wf_permit_setting') then
	drop procedure wf_permit_setting;
end if;

CREATE procedure wf_permit_setting(
	p_server_name varchar(32)
	, p_table_name varchar(64) default ''
	, p_operation varchar(32) default ''
	, p_table_field varchar(64) default ''
)
begin
	call wf_release_forbid (p_server_name, p_table_name, p_operation, p_fallback, p_field_name, 1);
end;


/*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/
/*                                                   */
/*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/
if exists (select 1 from sysprocedure where proc_name = 'wf_permit_session') then
	drop procedure wf_permit_session;
end if;

CREATE procedure wf_permit_session(
	p_server_name varchar(32)
	, p_table_name varchar(64) default ''
	, p_operation varchar(32) default ''
	, p_table_field varchar(64) default ''
)
begin
	call wf_release_forbid (p_server_name, p_table_name, p_operation, p_fallback, p_field_name, 0);
end;


