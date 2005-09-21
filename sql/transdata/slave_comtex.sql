--****************************************************************
--                               NEXT ID
--****************************************************************
if exists (select 1 from sysprocedure where proc_name = 'slave_nextid') then
	drop procedure slave_nextid;
end if;

create PROCEDURE slave_nextid(
		in table_name varchar(100)
		,out mid integer
	)
begin
  execute immediate 'select isnull(max(id), -1)+1 into mid from ' + table_name;
end;




if exists (select 1 from sysprocedure where proc_name = 'slave_nextnu') then
	drop procedure slave_nextnu;
end if;

create PROCEDURE slave_nextnu(
		  in p_table_name varchar(100)
		, out p_nu varchar(32)
		, in p_dat_field varchar(32) default null
	)
begin
	declare v_sql varchar(1000);

	if p_dat_field is null then
		set p_dat_field = 'dat';
	end if;

	set v_sql = 
		 ' select nu as r_nu from ' + p_table_name
		+'	where convert(varchar(4), ' + p_dat_field + ', 112) = convert(varchar(4), now(), 112)'
		+' order by'
		+'	if isnumeric(nu) = 1 then convert(integer, nu) else 0 endif desc'
	;
		
	begin
		declare c_product_variants cursor using v_sql;
	    
		open  c_product_variants;
		fetch c_product_variants into p_nu;
		set p_nu = convert(varchar(20), convert(integer, p_nu) + 1);
		close c_product_variants;
	end;

	if p_nu is null then
		set p_nu = '1';
	end if;

end;
	

--****************************************************************
--                      CURRENCY AND RATES
--****************************************************************

if exists (select 1 from sysprocedure where proc_name = 'slave_currency_rate') then
	drop procedure slave_currency_rate;
end if;

create PROCEDURE slave_currency_rate(
		out o_date varchar(20)
		,out o_rate float
		,in p_date varchar(20) default null
		,in p_id_cur integer default null
	)
begin
	
	declare v_date date;

	if p_date is null then
		set v_date = now();
	else 
		set v_date = convert(date, p_date);
	end if;

	set v_date = convert(date, convert(varchar(20), v_date, 112));

	if p_id_cur is null then
		select id into p_id_cur from currency where iso_code = 'UE';
	end if;

	select curse, mj.max_dat
	into o_rate, o_date
	from cur_rate cr
	join (
		select max(dat) max_dat, id_cur 
		from cur_rate m 
		where m.dat <= v_date and m.id_cur = p_id_cur group by id_cur
	) mj on mj.max_dat = cr.dat and mj.id_cur = cr.id_cur;

end;

/*
begin
	declare o_date date;
	declare o_curse float;
	call slave_currency_rate(now()-1,11,o_curse, o_date);
	select o_curse;
end;
*/

if exists (select 1 from sysprocedure where proc_name = 'slave_date_currency_rate') then
	drop function slave_date_currency_rate;
end if;

create function slave_date_currency_rate(
		in p_date date
		,in p_id_cur integer
	)
	returns float
begin
	declare o_date date;
	declare o_curse float;

	call slave_currency_rate(p_date, p_id_cur, o_curse, o_date);
	return o_curse;
end;


if exists (select 1 from sysprocedure where proc_name = 'slave_get_currency_rate') then
	drop function slave_get_currency_rate;
end if;

create function slave_get_currency_rate(
		in p_id_cur integer
	)
	returns float
begin
	return slave_date_currency_rate(now(), p_id_cur);
end;



--****************************************************************
--              INTEGRATION PRIOR/COMTEX
--****************************************************************

if exists (select 1 from sysprocedure where proc_name = 'get_standalone') then
	drop function get_standalone;
end if;


CREATE function get_standalone( 
	p_server char(50) default null
) returns integer
begin
	declare v_check varchar(23);

	set get_standalone = 1;
	select KIND into v_check from guides where id = 0;
	if v_check is null or v_check = '' or v_check = 0 then
		set get_standalone = 0;
	end if;
end;


if exists (select 1 from sysprocedure where proc_name = 'slave_get_standalone') then
	drop procedure slave_get_standalone;
end if;

CREATE procedure slave_get_standalone(
	out p_standalone integer
)
begin
	set p_standalone = get_standalone();
end;




if exists (select 1 from sysprocedure where proc_name = 'set_standalone') then
	drop function set_standalone;
end if;

CREATE function set_standalone(
	 p_standalone varchar(23)
) returns integer
begin
	set set_standalone = 1;
	update guides set kind = p_standalone where id = 0;

	exception when others then
		set set_standalone = 0;
end;



if exists (select 1 from sysprocedure where proc_name = 'slave_set_standalone') then
	drop procedure slave_set_standalone;
end if;

CREATE procedure slave_set_standalone(
	 out p_succes integer
	 , p_standalone varchar(23)
) 
begin
	set p_succes = set_standalone(p_standalone);
end;


