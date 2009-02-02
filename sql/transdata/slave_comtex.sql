

if exists (select '*' from sysprocedure where proc_name like 'change_mat_qty') then
	drop procedure change_mat_qty;
end if;

create 
	-- процедура вызывается для коррекции суммы по позиции в приходной накладной
	-- при изменении количества в интерфейсе stime
	-- 
procedure change_mat_qty (
	  in p_id_mat integer
	, in p_new_quant float
--	, in p_cena float default 0.0
)
begin
	declare v_kol1 float;
	select kol1 into v_kol1 from mat where id= p_id_mat;
	if abs(round(v_kol1, 3)) > 0.001 then
		update mat set 
			 summa       = summa       / kol1 * p_new_quant
			,summa_sale  = summa_sale  / kol1 * p_new_quant
			,summav      = summav      / kol1 * p_new_quant
			,summa_salev = summa_salev / kol1 * p_new_quant
			,kol1        = p_new_quant
			,kol2        = p_new_quant
			,kol3        = p_new_quant
		where id = p_id_mat;
	else
		update mat set 
			kol1        = p_new_quant
			,kol2        = p_new_quant
			,kol3        = p_new_quant
		where id = p_id_mat;
	end if;
end;


if exists (select '*' from sysprocedure where proc_name like 'order_import') then
	drop procedure order_import;
end if;

if exists (select '*' from sysprocedure where proc_name like 'change_id_guide') then
	drop procedure change_id_guide;
end if;


create procedure change_id_guide (
-- процедура должна вызываться при смене типа накладной с рублевой
-- на импортную или наоборот
-- Пересчет денежных значений по позициям должен осуществляться 
-- вне этой функции.
	  in p_id_jmat integer
	, in p_id_guide integer
	, in p_id_currency integer
	, in p_tp1 integer
	, in p_tp2 integer
	, in p_tp3 integer
	, in p_tp4 integer
)
begin

	declare out_cur_date varchar(20);
	declare v_rate float;


		-- текущий курс валюты
		call slave_currency_rate(out_cur_date, v_rate, null, p_id_currency);

		update jmat set 
			id_guide = p_id_guide
			, id_curr = p_id_currency 
			, tp1 = p_tp1
			, tp2 = p_tp2
			, tp3 = p_tp3
			, tp4 = p_tp4
			, curr = isnull(v_rate, 1.0)
		where 
			id = p_id_jmat;

end;




--****************************************************************
--                               NEXT ID
--****************************************************************
if exists (select 1 from sysprocedure where proc_name = 'slave_nextid') then
	drop procedure slave_nextid;
end if;

create PROCEDURE slave_nextid
-- Исходим из предположения, что в таблице лежит не следующий готовый id
-- а последний использованный. (Последние изменения для Komtex 9)
-- Наввание колонки осталось next_id, но смысл ее помеялся на last_id
(
	in table_name varchar(100)
	,out mid integer
)
begin
	select next_id into mid from inc_table where table_nm = table_name;
	if mid is null then
		execute immediate 'select isnull(max(id), 1) into mid from ' + table_name;
--		update inc_table set next_id = mid where current of dc;
		execute immediate 'insert into inc_table  (table_nm, next_id) select upper(''' + table_name + '''), mid';

		--call build_id_track_trigger(table_name);
	end if;
end;



--****************************************************************
--                               NEXT NU
--****************************************************************

if exists (select 1 from sysprocedure where proc_name = 'slave_nextnu') then
	drop procedure slave_nextnu;
end if;

create PROCEDURE slave_nextnu(
	  in p_table_name varchar(100)
	, out p_nu varchar(32)
	, in p_nu_old varchar(32)    default null
	, in p_dat_field varchar(32) default null
	, in p_dat varchar(32) default null
)
begin
	declare v_sql varchar(1000);

	if p_dat_field is null then
		set p_dat_field = 'dat';
	end if;

	if p_dat is null then
		set p_dat = now();
	end if;

	set v_sql = 
		 ' select isnull(max(convert(integer, nu)), 0) + 1 ' 
		+' into p_nu'
		+' from ' + p_table_name
		+'  where convert(varchar(4), ' + p_dat_field + ', 112) = convert(varchar(4), ''' + p_dat + ''', 112)'
		+'  and isnumeric(nu) = 1'
	;


	if p_nu_old is not null then
		set v_sql = v_sql
			+ ' and nu != ''' + convert(varchar(20), p_nu_old) + ''''
		;
	end if;

	message v_sql to client;
	execute immediate v_sql;
		
end;


if exists (select 1 from sysprocedure where proc_name = 'slave_renu_scet') then
	drop procedure slave_renu_scet;
end if;

create PROCEDURE slave_renu_scet(
	in p_id_jscet integer
)
begin
	declare new_nu integer;
	set new_nu = 1;
	for v_server_name as a dynamic scroll cursor for
		select id_jmat, nu from scet where id_jmat = p_id_jscet
		order by isnull(nu, 999999999)
		for update
	do
		
		update scet set nu = new_nu where current of a;
		set new_nu = new_nu + 1;

	end for
end;
	

if exists (select 1 from sysprocedure where proc_name = 'slave_move_uslug') then
	drop procedure slave_move_uslug;
end if;

create PROCEDURE slave_move_uslug(
	  in p_id_jscet integer
	, in p_id_jscet_new integer
	, in p_quant float
	, in p_id_inv integer
)
begin

	a:
	for v_server_name as a dynamic scroll cursor for
		select
			id, id_jmat
		from scet
		where
			id_jmat = p_id_jscet
		and id_inv =  p_id_inv
		and summa_salev = p_quant
		for update
	do
		update scet set id_jmat = p_id_jscet_new
		where current of a;

		leave a;
	end for;
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



if exists (select 1 from sysprocedure where proc_name = 'slave_legacy_purpose') then
	drop procedure slave_legacy_purpose;
end if;

create PROCEDURE slave_legacy_purpose(
		in purpose_name varchar(100)
		,in debit varchar(26)
		,in subdebit varchar(10)
		,in kredit varchar(26)
		,in subkredit varchar(10)
	)
begin
	declare v_id_d integer;
	declare v_id_c integer;

	select id into v_id_d from account d where d.sc = debit and d.sub_sc = subdebit;
	select id into v_id_c from account c where c.sc = kredit and c.sub_sc = subkredit;
	if v_id_d is null or v_id_c is null then
		return;
	end if;

	if not exists (select 1 from m_xoz where nm = purpose_name and id_accd = v_id_d and id_accc = v_id_d) then
		insert into m_xoz(nm, id_accd, id_accc)
		select purpose_name, v_id_d, v_id_c;
	end if;
end;


if exists (select 1 from sysprocedure where proc_name = 'slave_list_customer') then
	drop procedure slave_list_customer;
end if;
/*
create PROCEDURE slave_list_customer(
	  in p_id_vocnames integer
) result (
	 id       integer
	,FirmName varchar(98)
	,Inn      varchar(14)
	,Okonx    varchar(5)
	,Okpo     varchar(10)
	,Kpp      varchar(10)
	,Address  varchar(98)
	,phone    varchar(37)
)
begin
	select 
		 v.id      as id
		,v.nm      as FirmName 
		,p.inn     as inn
		,p.okonx   as okonx
		,p.okpo    as okpo
		,p.kpp     as kpp
		,v.address as address
	 	,v.phone   as phone
	from voc_names v
	join post p on p.id = v.id
	where 
		v.id = isnull(p_id_vocnames, v.id);
end;
*/


