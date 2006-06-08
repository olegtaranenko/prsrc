if exists (select '*' from sysprocedure where proc_name like 'wf_make_venture_income') then  
	drop function wf_make_venture_income;
end if;


create function wf_make_venture_income(
	p_numdoc integer
	, p_venture_id integer
) returns integer
-- устанавливает признак, на какое придриятие был осуществлен
-- приход
begin
	declare v_numdoc varchar(20);
	declare v_numext varchar(20);
	declare v_id_analytic_default integer;
	declare old_id_analytic integer;
	declare new_id_analytic integer;
	declare v_code_name varchar(30);
	declare v_id_jmat integer;
	declare v_activity_start date;
	declare v_xdate date;

	set wf_make_venture_income = 1;
	set v_numext = 255;
	set v_numdoc = p_numdoc;

	select d.id_jmat, i.id_analytic, v.id_analytic, s.id_analytic_default, v.activity_start, d.xdate
	into v_id_jmat, old_id_analytic, new_id_analytic, v_id_analytic_default, v_activity_start, v_xdate
	from sdocs d
	left join sdocsIncome i on i.numdoc = d.numdoc and i.numext = d.numext
	left join guideventure v on v.ventureId = p_venture_id
	join system s on 1=1
	where d.numdoc = v_numdoc and d.numext = v_numext;

	if v_activity_start > v_xdate then
		-- нельзя осуществить приход на предприятие до начала его работы
		set wf_make_venture_income = 0;
		return;
	end if;

	if v_id_jmat is not null then
		call update_remote('stime', 'jmat', 'id_code', isnull(new_id_analytic, 0), 'id = ' + convert(varchar(20), v_id_jmat));
	else
--		set wf_make_venture_income = 0;
	end if;
	-- Приходуем накладную на то или иное предприятие в зависимости от 
	-- кода аналитики
	-- для этого в таблицу sDocsIncome добавляем/удаляем строку со
	-- ссылкой на предприятие
	if new_id_analytic is null or new_id_analytic = v_id_analytic_default then
		delete from sdocsincome where numdoc = v_numdoc and numext = v_numext;
	else
		if old_id_analytic is null then
			insert into sdocsIncome (numdoc, numext, id_analytic, ventureId, id_jmat)
			values (v_numdoc, v_numext, new_id_analytic, p_venture_id, v_id_jmat);
	   	else 
	   		update sdocsIncome set id_analytic = new_id_analytic 
	   		where numdoc = v_numdoc and numext = v_numext;
		end if;


	end if;
	-- выставить признак того, что взаимозачеты необходимо пересчитать
	update sdocsventure dv set invalid = 1
	where v_xdate between dv.termFrom and dv.termTo
	and dv.cumulative_id is null;
end;
 

 


/* * * * * * * * * * * * * * * * * * * * * * * * * * *
 * Процедуры для получения себестоимости из комтеха
 * * * * * * * * * * * * * * * * * * * * * * * * * * */
if exists (select '*' from sysprocedure where proc_name like 'wf_get_comtex_cost') then  
	drop function wf_get_comtex_cost;
end if;



if exists (select '*' from sysprocedure where proc_name like 'wf_cost_bulk_change') then  
	drop function wf_cost_bulk_change;
end if;


create function wf_cost_bulk_change (
	p_klassid integer
	, p_cur_rate float default null
) returns integer
begin
	declare v_lvl integer;
	declare v_price_bulk_Id integer;
	declare v_comtex_cost float;
	declare v_timestamp datetime;
	declare v_cur_rate float;

	create table #tmp_klass(lvl integer, id integer);

	set v_lvl = 0;
	if p_klassid > 0 then
		insert into #tmp_klass (lvl, id) select 0, p_klassid;
	    
		branch: loop
			insert into #tmp_klass (lvl, id)
				select v_lvl + 1, k.klassId
				from sguideklass k
				join #tmp_klass t on t.id = k.parentKlassId and t.lvl = v_lvl;
	    
			if @@rowcount = 0 then
				leave branch;
			end if;
			set v_lvl = v_lvl + 1;
		end loop;
	else
		insert into #tmp_klass (lvl, id) 
		select 0, klassid
		from sguideklass
		where klassid != 0;
	end if;

	if p_cur_rate is not null then
		set v_cur_rate = p_cur_rate;
	else
		set v_cur_rate = system_currency_rate();
	end if;

	for v_table as b1 dynamic scroll cursor for
		select nomnom as r_nomnom, id_inv as r_id_inv
			, cost as r_prior_cost, perList as r_perlist
		from sguidenomenk n
		join #tmp_klass t on n.klassid = t.id
		where id_inv is not null
	do 
		call wf_calc_cost_stime(v_comtex_cost, r_id_inv);
		if v_comtex_cost > 0 then
			set v_comtex_cost = v_comtex_cost / v_cur_rate;
			if abs(round((v_comtex_cost - r_prior_cost), 2) ) > 0.01 then
				if v_price_bulk_Id is null then
					insert into sPriceBulkChange (guide_klass_id) values (p_klassid);
					set v_price_bulk_Id = @@identity;
				end if;
	    
				update sguidenomenk set cost = round(v_comtex_cost, 2) where nomnom = r_nomnom;
				-- триггером в этот момент добавляется запись в sPriceHistory
				select max(change_date) into v_timestamp from sPriceHistory where nomnom = r_nomnom;
				
				update sPriceHistory set bulk_id = v_price_bulk_id where change_date = v_timestamp and nomnom = r_nomnom;
	    
			end if;
		end if;
	end for;

	drop table #tmp_klass;

	return v_price_bulk_id;

end;

/* * * * * * * * * * * * * * * * * * * * * * * * * * *
 * Функция блокировки. (опять не работают)
 * * * * * * * * * * * * * * * * * * * * * * * * * * */

if exists (select '*' from sysprocedure where proc_name like 'bootstrap_blocking') then  
	drop procedure bootstrap_blocking;
end if;


create procedure bootstrap_blocking (
) 
begin
	call cre_block_var('blocks_inited');
	call cre_block_var('bulk_delete');
	call cre_block_var('supress_cum_update');
	call cre_block_var('supress_diary_update');

/*
	for v_table as b2 dynamic scroll cursor for
		select 'sdocs' as r_table union select 'sdmc' union select 'guidefirm' union select 'bayguidefirm'
	do 
		for v_server_name as a2 dynamic scroll cursor for
			select 
				srvname as r_server 
			from sys.sysservers s 
			join guideventure v on s.srvname = v.sysname
		do
			
			call cre_block_var(make_block_name(r_server, r_table));
		end for;
	end for;
*/
	
	for v_table as b1 dynamic scroll cursor for
		select 'jmat' as r_table union select 'mat' union select 'jscet' union select 'scet'
	do 
		for v_server_name as a1 dynamic scroll cursor for
			select 
				srvname as r_server 
			from sys.sysservers s 
			join guideventure v on s.srvname = v.sysname and v.standalone = 0
		do
			message 'call slave_cre_block_var_' + r_server + '(''' + make_block_name(@@servername, r_table) + ''')' to log;
			execute immediate 'call slave_cre_block_var_' + r_server + '(''' + make_block_name(@@servername, r_table) + ''')';

		end for;
	end for;


end;
	
	

if exists (select '*' from sysprocedure where proc_name like 'firstDayMonth') then  
	drop function firstDayMonth;
end if;


create function firstDayMonth (
	p_dt date
) returns date
begin
	set firstDayMonth = 
		convert(date, ymd(year(p_dt), month(p_dt), 1));
end;



if exists (select '*' from sysprocedure where proc_name like 'lastDayMonth') then  
	drop function lastDayMonth;
end if;


create function lastDayMonth (
	p_dt date
) returns date
begin
	set lastDayMonth = 
			convert(date, ymd(year(p_dt), 1 + month(p_dt), 1) - 1)
end;


-------------------------------------------------------------------------
--------------       sDmcVenture triggers         ----------------------
-------------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_cumulative_del' and tname = 'sDmcVenture') then 
	drop trigger sDmcVenture.wf_cumulative_del;
end if;

create TRIGGER wf_cumulative_del before delete order 1 on
sDmcVenture
referencing old as old_name
for each row
when (exists (select 1 from sdocsventure n where n.id = old_name.sdv_id and n.cumulative_id is null))
begin
	declare v_cumulative_id integer;
	declare v_dmc_id integer;
	declare v_quant float;
	declare no_echo integer;

	select cumulative_id 
	into v_cumulative_id
	from sDocsVenture where id = old_name.sdv_id;

		-- удаляем позицию из сводной => нужно удалить все зачеты 
		-- из дневных накладных

--		execute immediate 'create variable @bulk_delete integer';
		set @bulk_delete = 1;
		select @bulk_delete into no_echo;

		delete from sdmcVenture diary
		from  
			  sDocsVenture diary_doc
		where diary.nomnom = old_name.nomnom
			and diary.sdv_id = diary_doc.id
			and diary_doc.cumulative_id = old_name.sdv_id
		;

		set @bulk_delete = 0;
--		execute immediate 'drop variable @bulk_delete';
end;

if exists (select 1 from systriggers where trigname = 'wf_diary_del' and tname = 'sDmcVenture') then 
	drop trigger sDmcVenture.wf_diary_del;
end if;

create TRIGGER wf_diary_del before delete order 2 on
sDmcVenture
referencing old as old_name
for each row
when (exists (select 1 from sdocsventure n where n.id = old_name.sdv_id and n.cumulative_id is not null))
begin
	declare v_cumulative_id integer;
	declare v_dmc_id integer;
	declare v_quant float;
	declare no_echo integer;

	select cumulative_id 
	into v_cumulative_id
	from sDocsVenture where id = old_name.sdv_id;

		-- скорректировать количество по позиции в сводной накладной

		
  	  	begin
			select @bulk_delete into no_echo; 
		exception 
			when other then
				set no_echo = 0;
		end;
	    
		if no_echo = 1 then
			return;
		end if;
		
		
		select 
			m.id
		into v_dmc_id
		from sDmcVenture m
		join sDocsVenture n on m.sdv_id = n.id 
			and n.id = v_cumulative_id
		where m.nomnom = old_name.nomnom
		;

		if v_dmc_id is not null then
			update sDmcVenture set quant = quant - old_name.quant 
			where id = v_dmc_id;
			select quant into v_quant from sDmcVenture where id = v_dmc_id;

			if round(v_quant, 3) < 0.001 then
				delete from sDmcVenture where id = v_dmc_id;
			end if;
			
		end if;
end;



if exists (select 1 from systriggers where trigname = 'wf_cumulative_upd' and tname = 'sDmcVenture') then 
	drop trigger sDmcVenture.wf_cumulative_upd;
end if;


create TRIGGER wf_cumulative_upd before update order 1 on
sDmcVenture
referencing old as old_name new as new_name
for each row
when (exists (select 1 from sdocsventure n where n.id = old_name.sdv_id and n.cumulative_id is null))
begin
	declare v_cumulative_id integer;
	declare v_dmc_id integer;
	declare v_quant float;
	declare v_ratio float;
	declare no_echo integer;
	declare v_cum_quant float;
	declare old_cum_costed float;
	declare old_cum_quant float;
	declare v_cumulative_total float;


  	  	begin
			select @supress_cum_update into no_echo; 
		exception 
			when other then
				set no_echo = 0;
		end;
	    
		if no_echo = 1 then
			return;
		end if;
		

		if update(costed) then
			-- пропорционально изменить сумму в дневных накладных
			-- так, чтобы сумма сводной билась с дневными
--			execute immediate 'create variable @supress_diary_update integer';
			set @supress_diary_update = 1;
			select @supress_diary_update into no_echo;

		    --message 'wf_cumulative_upd::old_name.costed = ', old_name.costed to client;
		    --message 'wf_cumulative_upd::new_name.costed = ', new_name.costed to client;
			if old_name.costed = 0 then
				update sdmcVenture diary set costed = new_name.costed
				from  
					sDocsVenture diary_doc
				where diary.nomnom = old_name.nomnom
					and diary.sdv_id = diary_doc.id
					and diary_doc.cumulative_id = old_name.sdv_id 
					;
			else 

				set v_ratio = new_name.quant * new_name.costed / (old_name.quant * old_name.costed);
		    
				update sdmcVenture diary set costed = costed * v_ratio
				from  
					 sDocsVenture diary_doc
				where diary.nomnom = old_name.nomnom
					and diary.sdv_id = diary_doc.id
					and diary_doc.cumulative_id = old_name.sdv_id 
					;
			end if;
	    
			set @supress_diary_update = 0;
--			execute immediate 'drop variable @supress_diary_update';
	    
		end if;
end;

if exists (select 1 from systriggers where trigname = 'wf_diary_upd' and tname = 'sDmcVenture') then 
	drop trigger sDmcVenture.wf_diary_upd;
end if;


create TRIGGER wf_diary_upd before update order 2 on
sDmcVenture
referencing old as old_name new as new_name
for each row
when (exists (select 1 from sdocsventure n where n.id = old_name.sdv_id and n.cumulative_id is not null))
begin
	declare v_cumulative_id integer;
	declare v_dmc_id integer;
	declare v_quant float;
	declare v_ratio float;
	declare no_echo integer;
	declare v_cum_quant float;
	declare old_cum_costed float;
	declare old_cum_quant float;
	declare v_cumulative_total float;



		begin
			select @supress_diary_update into no_echo; 
		exception 
			when other then
				set no_echo = 0;
		end;
	    
		if no_echo = 1 then
			return;
		end if;
		
--		    message 'wf_diary_upd::old_name.costed = ', old_name.costed to client;
--		    message 'wf_diary_upd::new_name.costed = ', new_name.costed to client;
	select cumulative_id 
	into v_cumulative_id
	from sDocsVenture where id = old_name.sdv_id;

		-- скорректировать количество по позиции в сводной накладной
		if update(quant) or update(costed) then

			select m.id, m.quant, m.costed
			into v_dmc_id, old_cum_quant, old_cum_costed
			from sDmcVenture m
			join sDocsVenture n on m.sdv_id = n.id and n.id = v_cumulative_id
			where m.nomnom = old_name.nomnom;

			if v_dmc_id is not null then

--				execute immediate 'create variable @supress_cum_update integer';
				set @supress_cum_update = 1;
				select @supress_cum_update into no_echo;

		    	set v_cum_quant = old_cum_quant - old_name.quant + new_name.quant;
		        set v_cumulative_total = (old_cum_quant * old_cum_costed) 
		        		- (old_name.quant * old_name.costed) 
		        		+ (new_name.quant * new_name.costed)
		        ;
	    
				update sDmcVenture 
					set costed = v_cumulative_total / v_cum_quant
					, quant = v_cum_quant
				where id = v_dmc_id;

				set @supress_cum_update = 0;
--				execute immediate 'drop variable @supress_cum_update';

			end if;
		end if;
end;



if exists (select 1 from systriggers where trigname = 'wf_cumulative_add' and tname = 'sDmcVenture') then 
	drop trigger sDmcVenture.wf_cumulative_add;
end if;

create TRIGGER wf_cumulative_add before insert order 1 on
sDmcVenture
referencing new as new_name
for each row
begin
	declare v_cumulative_id integer;
	declare v_dmc_id integer;
	declare no_echo integer;



	select cumulative_id 
	into v_cumulative_id
	from sDocsVenture where id = new_name.sdv_id;
	--message 'v_cumulative_id = ', v_cumulative_Id to client;


	if v_cumulative_id is not null then
		-- добавить (или проапдейтить) позицию в сводной накладной
		select 
			m.id
		into v_dmc_id
		from sDmcVenture m
		join sDocsVenture n on m.sdv_id = n.id 
			and n.id = v_cumulative_id
		where m.nomnom = new_name.nomnom
		;
		--message 'v_dmc_id = ', v_dmc_id to client;

		if v_dmc_id is null then
			insert into sDmcVenture (
				sdv_id, nomnom, quant, costed
			) values (
				v_cumulative_id
				, new_name.nomnom
				, new_name.quant
				, new_name.costed
			);
		else
--			execute immediate 'create variable @supress_cum_update integer';
			set @supress_cum_update = 1;
			select @supress_cum_update into no_echo;

			update sDmcVenture set quant = quant + new_name.quant 
			where id = v_dmc_id;

			set @supress_cum_update = 0;
--			execute immediate 'drop variable @supress_cum_update';
		end if;
			
	end if;
end;
-------------------------------------------------------------------------
--------------       end of sDmcVenture triggers         ---------------
-------------------------------------------------------------------------





-------------------------------------------------------------------------
--------------       sDocsVenture triggers         ----------------------
-------------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_cumulative_del' and tname = 'sDocsVenture') then 
	drop trigger sDocsVenture.wf_cumulative_del;
end if;

create TRIGGER wf_cumulative_del before delete order 1 on
sDocsVenture
referencing old as old_name
for each row
when (old_name.cumulative_id is null)
begin
end;



if exists (select 1 from systriggers where trigname = 'wf_cumulative_upd' and tname = 'sDocsVenture') then 
	drop trigger sDocsVenture.wf_cumulative_upd;
end if;

create TRIGGER wf_cumulative_upd before update order 1 on
sDocsVenture
referencing old as old_name new as new_name
for each row
when (old_name.cumulative_id is null)
begin
end;



if exists (select 1 from systriggers where trigname = 'wf_cumulative_add' and tname = 'sDocsVenture') then 
	drop trigger sDocsVenture.wf_cumulative_add;
end if;

create TRIGGER wf_cumulative_add before insert order 1 on
sDocsVenture
referencing new as new_name
for each row
when (new_name.cumulative_id is null)
begin
end;

if exists (select 1 from systriggers where trigname = 'wf_diary_del' and tname = 'sDocsVenture') then 
	drop trigger sDocsVenture.wf_diary_del;
end if;

create TRIGGER wf_diary_del before delete order 2 on
sDocsVenture
referencing old as old_name
for each row
when (old_name.cumulative_id is not null)
begin
	delete from sdmcVenture cum
	from sdmcVenture diary
	where cum.sdv_id = old_name.cumulative_id 
		and diary.sdv_id = old_name.id
		and cum.nomnom = diary.nomnom;
end;



if exists (select 1 from systriggers where trigname = 'wf_diary_upd' and tname = 'sDocsVenture') then 
	drop trigger sDocsVenture.wf_diary_upd;
end if;

create TRIGGER wf_diary_upd before update order 2 on
sDocsVenture
referencing old as old_name new as new_name
for each row
when (old_name.cumulative_id is not null)
begin
end;



if exists (select 1 from systriggers where trigname = 'wf_diary_add' and tname = 'sDocsVenture') then 
	drop trigger sDocsVenture.wf_diary_add;
end if;

create TRIGGER wf_diary_add before insert order 2 on
sDocsVenture
referencing new as new_name
for each row
when (new_name.cumulative_id is not null)
begin
	declare v_cumulative_id integer;
	-- предполагаем, что дневной взаимозачет вставляем с id = 0
	if new_name.cumulative_id = 0 then
		select id into v_cumulative_id from sDocsVenture
		where 
				termFrom = isnull(new_name.termFrom, firstDayMonth(new_name.nDate))
			and termTo   = isnull(new_name.termTo, lastDayMonth(new_name.nDate))
			and srcVentureId = new_name.srcVentureId
			and dstVentureId = new_name.dstVentureId
			and cumulative_id is null;

		if v_cumulative_Id is null then
			insert into sdocsventure (
				termFrom
				, termTo
				, srcVentureId
				, dstVentureId
				, cumulative_id
				, nDate
				, procent
			) values (
				  isnull(new_name.termFrom, firstDayMonth(new_name.nDate))
				, isnull(new_name.termTo, lastDayMonth(new_name.nDate))
				, new_name.srcVentureId
				, new_name.dstVentureId
				, null
				, firstDayMonth(new_name.nDate)
				, new_name.procent
			);
			set v_cumulative_id = @@identity;
		end if;
		set new_name.cumulative_id = v_cumulative_id;
	end if;

	if new_name.termFrom is null then
		set new_name.termFrom = firstDayMonth(new_name.nDate);
	end if;
	if new_name.termTo is null then
		set new_name.termTo = lastDayMonth(new_name.nDate);
	end if;
end;

-------------------------------------------------------------------------
--------------       end of sDocsVenture triggers         ---------------
-------------------------------------------------------------------------

if exists (select '*' from sysprocedure where proc_name like 'ivo_validate') then  
	drop procedure ivo_validate;
end if;


create procedure ivo_validate (
	  p_procentOver float default null
) 
begin
	declare v_invalidate integer;
	declare v_term_min date;
	declare v_term_max date;

	set v_invalidate = 0;
	for ivo_c as ivo dynamic scroll cursor for
		select 
			id_jmat as r_id_jmat 
			, n.id as r_ivo_id
			, d.id_analytic as r_id_analytic
			, termFrom as r_term_start
			, termTo as r_term_end
			, nDate as r_nDate
			, isnull(invalid, 0) as r_invalid
			, n.procent as r_ivo_procent
		from sDocsVenture n
		join guideVenture s on s.ventureId = n.srcVentureId
		join guideVenture d on d.ventureId = n.dstVentureId
--		where isnull(n.invalid, 0) = 1
		order by n.ndate
	do
		if r_invalid = 1 then
			set v_invalidate = 1;
			update sdocsventure set invalid = 0 where id = r_ivo_id;
		end if;
		if v_invalidate = 1 then
			call ivo_comtex_delete(r_ivo_id);
			delete from sdocsventure where cumulative_Id = r_ivo_id;
			--delete from sdmcventure where sdv_id = r_ivo_id;
			if r_term_start <= isnull(v_term_min, '20000101') then
				set v_term_min = r_term_start;
			end if;
			if r_term_end >= isnull(v_term_max, '21000101') then
				set v_term_max = r_term_end;
			end if;
		end if;
	end for;
	call fill_venture_order(
		p_procentOver
		, v_term_min
		, v_term_max
	);
end;


if exists (select '*' from sysprocedure where proc_name like 'ivo_comtex_delete') then  
	drop procedure ivo_comtex_delete;
end if;


create procedure ivo_comtex_delete (
	 p_ivo_id integer
) 
begin

	for ivo_c as ivo dynamic scroll cursor for
		select 
			id_jmat as r_id_jmat 
			, d.id_analytic as r_id_analytic
			, termFrom as r_term_start
			, termTo as r_term_end
			, nDate as r_nDate
		from sDocsVenture n
		join guideVenture s on s.ventureId = n.srcVentureId
		join guideVenture d on d.ventureId = n.dstVentureId
		where id = p_ivo_id
	do
		if r_id_jmat is not null then
			call block_remote('stime', @@servername, 'jmat');
			call block_remote('stime', @@servername, 'mat');

			call delete_remote('stime', 'jmat', 'id = '+ convert(varchar(20), r_id_jmat));

			call unblock_remote('stime', @@servername, 'mat');
			call unblock_remote('stime', @@servername, 'jmat');

			update sdocsVenture set id_jmat = null where id = p_ivo_id;
		end if;
	end for;
	
end;




if exists (select '*' from sysprocedure where proc_name like 'ivo_to_comtex') then  
	drop procedure ivo_to_comtex;
end if;


create procedure ivo_to_comtex (
	 p_ivo_id integer
) 
begin
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_jmat_nu varchar(20);
	declare v_mat_nu integer;
	declare v_currency_rate float;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_id_source integer;
	declare v_id_dest integer;
	declare v_osn varchar(100);
	declare v_id_guide_jmat integer;
	declare v_folder_id integer;
	declare char_id_jmat varchar(20);



	for ivo_c as ivo dynamic scroll cursor for


		select 
			id_jmat as r_id_jmat 
			, d.id_analytic as r_id_analytic
			, termFrom as r_term_start
			, termTo as r_term_end
			, nDate as r_nDate
			, s.rusAbbrev as r_srcAbbrev
			, d.rusAbbrev as r_dstAbbrev
		from sDocsVenture n
		join guideVenture s on s.ventureId = n.srcVentureId
		join guideVenture d on d.ventureId = n.dstVentureId
		where id = p_ivo_id
	do
		if r_id_jmat is not null then
			-- проверим, не удалена ли она в Комехе?
			set char_id_jmat = select_remote('stime', 'jmat', 'id', 'id = ' + convert(varchar(20), r_id_jmat));
			if char_id_jmat is not null then
				-- удалить деталировку
				call delete_remote('stime', 'mat', 'jmat = '+ char_id_jmat);
			end if;
		else 
			set r_id_jmat = get_nextid('jmat');
			set char_id_jmat = null;
		end if;

		if char_id_jmat is null then
			-- вставить запись в журнал накладных.

			set v_id_guide_jmat = 1220;
	    
			set v_id_currency = system_currency();
			call slave_currency_rate_stime(v_datev, v_currency_rate);
			set v_jmat_nu = convert(varchar(20), p_ivo_id);
			--message 'v_jmat_nu = ', v_jmat_nu to client;
--			set v_jmat_nu = nextnu
	    
			select id_voc_names 
			into v_id_source 
			from sguidesource 
			where sourceid = -1001;
	    
			set v_id_dest = v_id_source;
			set v_osn = 'Вз/зачет с '
				+ convert(varchar(20), r_term_start, 104)
				+ ' по '+ convert(varchar(20), r_term_end, 104)
				+ '. Передача материалов от ' + r_srcAbbrev + ' => ' + r_dstAbbrev
			;
		    
			call block_remote('stime', @@servername, 'jmat');
			call wf_insert_jmat (
				'stime'
				,v_id_guide_jmat
				,r_id_jmat
				,r_nDate
				,v_jmat_nu
				,v_osn
				,v_id_currency
				,v_datev
				,v_currency_rate
				,v_id_source
				,v_id_dest
				,0
				,r_id_analytic
			);
			call unblock_remote('stime', @@servername, 'jmat');
	    
			update sDocsVenture d set 
				id_jmat = r_id_jmat
			where 
				id = p_ivo_id;
			;
		end if;




		set v_mat_nu = 1;
		-- добавить деталировку по сводной накладной
		for anomnom as c_nomnom dynamic scroll cursor for
			select 
				  quant as r_qty
				, m.nomnom as r_nomnom
				, perlist as r_perList
				, m.costed as r_cost
				, k.id_inv as r_id_inv
			from sdocsventure d
			join sdmcventure m on m.sdv_id = d.id
			join sguidenomenk k on k.nomnom = m.nomnom
			where 
				d.id_jmat = r_id_jmat
		do
			
			--message r_nomnom to client;
			call block_remote('stime', @@servername, 'mat');
			set v_id_mat = wf_insert_mat (
				'stime'
				,null
				,r_id_jmat
				,r_id_inv
				,v_mat_nu
				,r_qty 
				,r_cost
				,v_currency_rate
				,v_id_source
				,v_id_dest
				,r_perList
			);
			call unblock_remote('stime', @@servername, 'mat');
	    
	        update sdmcventure m set m.id_mat = v_id_mat 
	        where   r_id_jmat = m.sdv_id
				and m.nomnom = r_nomnom
			;

			set v_mat_nu = v_mat_nu + 1;

		end for;
	end for;
end;



if exists (select '*' from sysprocedure where proc_name like 'wf_put_ivo_nomnom') then  
	drop function wf_put_ivo_nomnom;
end if;


create function wf_put_ivo_nomnom (
	  p_target_date date
	, p_nomnom varchar(50)
	, p_qty    float
	, p_procent float
	, p_srcVentureid integer
	, p_dstVentureId integer
	, p_term_start date    default null
	, p_term_end date      default null
) returns integer
begin
	declare v_ndate date;
	declare v_nomnom varchar(20);
	declare v_costed float;
	declare v_perList float;
	declare v_procent float;
	declare chk_forward_income_qty float;

	select d.id, m.nomnom, isnull(m.quant, 0), procent
	into wf_put_ivo_nomnom, v_nomnom, chk_forward_income_qty, v_procent
	from sdocsventure d
	left join sdmcventure m on m.sdv_id = d.id and m.nomnom = p_nomnom
	where d.nDate = p_target_date 
		and srcVentureId = p_srcVentureId
		and dstVentureId = p_dstVentureId
 		and d.cumulative_id is not null
	;

	if wf_put_ivo_nomnom is null then
		insert into sDocsVenture (nDate, srcVentureId, dstVentureId, procent, termFrom, termTo)
		values (p_target_date, p_srcVentureId, p_dstVentureId, isnull(p_procent, v_procent), p_term_start, p_term_end);
		set wf_put_ivo_nomnom = @@identity;
	end if;

	select cena1, perList into v_costed, v_perList from sguidenomenk where nomnom = p_nomnom;


	if v_nomnom is null then
		insert into sDmcVenture(sdv_id, nomnom, quant, costed)
		select wf_put_ivo_nomnom, p_nomnom, p_qty * v_perList, v_costed * (1 + p_procent / 100);
	else 

		update sDmcVenture set quant = quant + p_qty * v_perList
		where sdv_id = wf_put_ivo_nomnom and nomnom = p_nomnom;
	end if;

end;



if exists (select '*' from sysprocedure where proc_name like 'fill_venture_order') then  
	drop procedure fill_venture_order;
end if;

create procedure fill_venture_order (
	  p_procentOver float
	, p_term_start date    default null
	, p_term_end date      default null
	, p_nomnom varchar(50) default null
)

begin
	declare total_rest float;
	declare rest1 float;
	declare rest2 float;
	declare v_defaultVentureId integer;
	declare v integer;
	declare cnt integer;
	declare vo_summa float;
	declare ivo_id integer;


	
	select v.ventureId into v_defaultVentureId 
	from guideventure v 
	join system s on s.id_analytic_default = v.id_analytic;
--	message v_defaultVentureId to client;


	create table #vntRest (ventureId integer, rest float);
		
	insert into #vntRest (ventureId, rest)
	select ventureId, 0.0
	from guideVenture where id_analytic is not null;

	for cur_nom as cn sensitive cursor for
		select nomnom as r_nomnom, perList as r_perList 
		from sguidenomenk n 
		where nomnom = isnull(p_nomnom, nomnom)
	do
	
		update #vntRest set rest = 0.00;
		
		nomnom_loop:
		for cur_history as his sensitive cursor for
			select 
				if destId <= -1001 then 2 else 3 endif 
    			        as sec_sort
				, convert(date, xDate) as r_nDate
				, n.sourid 
				, n.destid 
				, quant/k.perlist as r_qty
				, 	if (n.sourid <= -1001 and n.destid <= -1001) then 
						0 
					else 
						if n.destid <= -1001 then 
							1
						else
							-1
						endif
					endif as 
				r_activeOper
				, if (n.sourid <= -1001 and n.destid <= -1001) then 
						null 
					else 
						if n.destid <= -1001 then 
							isnull(i.ventureid, v.ventureid) 
						else 
							isnull(
								isnull(
									isnull(o.ventureid, bo.ventureid)
									, if substring(isnull(o.invoice, bo.invoice), 1, 2) = '55' then 2 else 1 endif
								), v.ventureid
							) 
						endif
					endif 
				as r_ventureid 
				, 0 as r_destVentureId
				, convert(varchar(20), n.numdoc) + '/' + convert(varchar(20),n.numext) as r_numdoc
			from sdocs n
				left join sdocsincome i on i.numdoc = n.numdoc and i.numext = n.numext
				join sdmc m on n.numdoc = m.numdoc 
						and n.numext = m.numext 
						and isnull(i.nomnom, m.nomnom) = m.nomnom
				join sguidenomenk k on k.nomnom = m.nomnom
			    join sguidesource s on s.sourceId = n.sourId
				join sguidesource d on d.sourceId = n.destId
				join system sys on 1 = 1
				join guideventure v on v.id_analytic = sys.id_analytic_default
				left join orders o on o.numorder = n.numdoc
				left join bayorders bo on bo.numorder = n.numdoc
			where
					m.nomnom = r_nomnom
				and convert(date, n.xDate) <= isnull(p_term_end, convert(date, n.xDate))
						union
			select 
				  1 as sec_sort 
				, n.nDate as r_nDate
				, null as sourId, null as destId
				, m.quant / k.perlist as r_qty
				, 0 as r_activeOper
				, srcVentureId as r_ventureId
				, dstVentureId as r_destVentureId
				, convert(varchar(20), n.id) as r_numdoc
			from sdmcventure m
			join sdocsventure n on m.sdv_id = n.id and n.cumulative_id is not null
			join sguidenomenk k on k.nomnom = m.nomnom
			where m.nomnom = r_nomnom
				and n.nDate <= isnull(p_term_end, n.nDate)
			order by 2, 1
		do
--			if r_nDate > isnull(p_term_end, r_nDate) then
--				leave nomnom_loop;
--			end if;

			if r_destVentureId > 0 then
				update #vntRest set rest = rest + r_qty where ventureId = r_destVentureId;
				update #vntRest set rest = rest - r_qty where ventureId = r_ventureId;
			else
				if not exists (select 1 from #vntRest where ventureId = r_ventureId) then
					set r_ventureId = v_defaultVentureId;
				end if; 
				update #vntRest set rest = rest + r_qty * r_activeOper where ventureId = r_ventureId;
			end if;
			
			if exists (select 1 from #vntRest where rest < 0)
				and r_nDate >= isnull(p_term_start, r_nDate)
--				and r_nDate <= isnull(p_term_end, r_nDate)
			then
				select sum(rest) into total_rest from #vntRest;
				if abs(round(total_rest, 3)) >= 0 then 
					-- найти где образовался минус, чтобы его компенсировать за счет того, у кого плюс
					compensate:
					for dst_vent as dv sensitive cursor for
						select rest as dv_rest, ventureId as dv_dstVentureId 
						from #vntRest where round(rest, 3) < 0
					do
						--message 'r_ndate = ', r_ndate, '    dv_rest = ', round(dv_rest, 2), '     total_rest =', round(total_rest, 2), '     r_ventureId = ', r_ventureId to client;
						set vo_summa = truncnum( abs(dv_rest) + 0.999999, 0);

						--message 'vo_summa = ', vo_summa to client;
					    for src_vent as sv sensitive cursor for
						select rest as sv_rest, ventureId as sv_srcVentureId 
							from #vntRest vr
							where 
								round(rest - vo_summa, 3) >= 0
								and vr.ventureId != dv_dstVentureId
						do
							-- Проверка на "паразитное" добавление после прихода товара на 
							-- фирму у которой не было отрицательного остатка, а у другой фирме
							-- этот отрицательный остаток был. Из-за того, что приход проходит
							-- всегда первым, взаимозачет увеличивается каждый раз при запуске процедуры

							if sec_sort <> 1 then
								--message 'sv_rest = ', sv_rest to client;
								set ivo_id = wf_put_ivo_nomnom (
									  convert(date, r_nDate)
									, r_nomnom
									, vo_summa
									, p_procentOver
									, sv_srcVentureId
									, dv_dstVentureId
									, p_term_start
									, p_term_end
								);
								update #vntRest set rest = rest + vo_summa where ventureId = dv_dstVentureId;
								update #vntRest set rest = rest - vo_summa where ventureId = sv_srcVentureId;
								leave compensate;
							end if;
						end for;
					end for;
				end if;
			end if;
/*			
			select rest into rest1 from #vntRest where ventureId = 1;
			select rest into rest2 from #vntRest where ventureId = 2;
			message r_numDoc, '        ', r_nDate,'      ', r_qty, '      ', round(rest1 + rest2, 2), '      ', round(rest1, 2), '      ', round(rest2, 2) 
			, '      ', r_ventureId
			, '      ', r_destventureId
			to client;
*/
/*
select d1.ndate , m1.nomnom
from sdmcventure m1
join sdocsventure d1 on m1.sdv_id = d1.id 
join sdmcventure m2 on m2.nomnom = m1.nomnom
join sdocsventure d2 on d2.ndate = d1.ndate and m2.sdv_id = d2.id and d1.srcventureId = d2.dstventureId
*/
		end for;
			
	end for;
	drop table #vntRest;
end;
		
	

if exists (select '*' from sysprocedure where proc_name like 'wf_make_invnm') then  
	drop function wf_make_invnm;
end if;


create function wf_make_invnm (
/* * * * * * * * * * * * * * * * * * * * * * * * * * *
 * Функция wf_make_invnm используется для получения
 * такого названия НЕВАРИАНТНОГО ИЗДЕЛИЯ или НОМЕНКЛАТУРЫ,
 * как оно будет выглядеть в базах Комтех.
 * В приоре это название не хранится в базе, а составляется
 * динамически из Cod, NomName, Size при показе в гриде.
 * В Комтехе это приходится прописывать жестко, как название 
 * позиции номенклатуры
 * * * * * * * * * * * * * * * * * * * * * * * * * * */
	  p_nomname varchar(50) default null
	, p_size varchar(30) default null
	, p_cod varchar(20) default null
) returns varchar(150)
begin
	    if (p_cod is not null and char_length(p_cod) > 0) then
	    	set wf_make_invnm =
	    		+ p_cod + ' ';
	    end if;

	    set wf_make_invnm = wf_make_invnm + p_nomname;
	    if (p_size is not null and char_length(p_size) > 0) then
	    	set wf_make_invnm = wf_make_invnm 
	    		+ ' ' + p_size;
	    end if;
end;

if exists (select '*' from sysprocedure where proc_name like 'wf_make_variant_nm') then  
	drop function wf_make_variant_nm;
end if;

create function wf_make_variant_nm (
/* * * * * * * * * * * * * * * * * * * * * * * * * * *
 * Функция wf_make_variant_nm используется для получения
 * такого названия ВАРИАНТНОГО ИЗДЕЛИЯ, 
 * как оно будет выглядеть в базах Комтех.
 * * * * * * * * * * * * * * * * * * * * * * * * * * */
	  p_nomname varchar(50) default null
	, p_size varchar(30) default null
	, p_cod varchar(20) default null
	, p_xprext varchar(20) default null
) returns varchar(150)
begin
	set wf_make_variant_nm = wf_make_invnm(p_nomName, p_size
		, convert(varchar(2), p_xprext) + '/' + p_cod
	);
	
end;


if exists (select '*' from sysprocedure where proc_name like 'wf_retrieve_bill_company') then  
	drop function wf_retrieve_bill_company;
end if;

create function wf_retrieve_bill_company (
	  p_id_bill integer
	, p_ventureName varchar(50)
) returns varchar(150)
begin
	declare v_serverName varchar(20);

	select sysname into v_serverName 
	from GuideVenture where ventureName = p_ventureName;
    
    --message 'sysname = ', v_serverName  to client;

		set wf_retrieve_bill_company = select_remote(
			v_serverName
			, 'voc_names'
			, 'nm'
			, 'id = ' + convert( varchar(20), p_id_bill)
		);
end;



if exists (select '*' from sysprocedure where proc_name like 'wf_check_jscet_split') then  
	drop function wf_check_jscet_split;
end if;

// возвращает id счета (бухгалтерского) из которого будет удаляться
create function wf_check_jscet_split (
	p_numorder integer            // заказ, которому меняем номер счета
) returns integer
begin
	declare remoteServerOld varchar(32);
	declare varchar_id varchar(20);
	declare v_invoice varchar(10);
	declare f_exists integer;

	// аттрибуты заказа который может быть слит с други
	// (тот у которого руками меняем номер счета)
	declare old_invoice varchar(10);      
	declare old_ventureId integer;
	declare old_id_jscet integer;
	declare old_invCode varchar(20);
	declare old_server varchar(20);
	declare old_numorder integer;

	set wf_check_jscet_split = null;

	select numorder, invoice, id_jscet, o.ventureId, v.invCode, v.sysname 
	into old_numorder, old_invoice, old_id_jscet, old_ventureId, old_invCode, old_server
	from orders o
		join guideventure v on v.ventureId = o.ventureId
	where numorder = p_numorder;

	if old_ventureId is null then
		return;
	end if;

	select count(*)
	into wf_check_jscet_split
	from orders o
    where o.invoice = old_invoice
	and o.numorder != old_numorder
	and isnull(o.shipped, 0) = 0;

end;


-------------------------------------
-------------------------------------
-------------------------------------
if exists (select '*' from sysprocedure where proc_name like 'wf_jscet_handle') then 
	drop function wf_jscet_handle;
end if;

// id бухгалтерского счета для заказа
create function wf_jscet_handle (
	// заказ, который должен быть выделен в отдельный счет
	  p_numorder integer			
	, in p_id_jscet_new integer default null
) returns integer
begin
	// аттрибуты заказа который может быть разделен на два
	declare old_invoice varchar(10);
	declare old_ventureId integer;
	declare old_id_jscet integer;
	declare old_invCode varchar(20);
	declare old_server varchar(20);
	declare v_nu_jscet varchar(50);
	declare v_id_jscet integer;
	declare v_id_dest integer;

	--message 'p_numorder = ', p_numorder to client;
	--message 'p_id_jscet_new = ', p_id_jscet_new to client;

	select invoice, id_jscet, o.ventureId, v.invCode, v.sysname, f.id_voc_names 
	into old_invoice, old_id_jscet, old_ventureId, old_invCode, old_server, v_id_dest
	from orders o
		join guideventure v on v.ventureId = o.ventureId
		join guidefirms f on f.firmid = o.firmid
	where numorder = p_numorder;

	if old_ventureId is null then
		return;
	end if;

	if p_id_jscet_new is not null then
		// делаем перемещение пре
		set v_id_jscet = p_id_jscet_new;
		set v_nu_jscet = select_remote (old_server, 'jscet', 'nu', 'id = ' + convert(varchar(20), p_id_jscet_new));
//		set out_invoice = old_invCode + convert(varchar(20), v_nu_jscet);
	else
		// выделение заказа в отдельный счет
		call put_jscet (v_id_jscet, v_nu_jscet, old_server, p_numOrder, v_id_dest, old_invoice);
	end if;

	update orders set id_jscet = v_id_jscet where numOrder = p_numorder;
	update orders set invoice = old_invCode + convert(varchar(20), v_nu_jscet) where numOrder = p_numorder;

	// Нужно выделить только те детали счета, которые относятся 
	// к заказу и перенести их в новый счет
	--message ' old_server = ', old_server to client;
	--message ' v_id_jscet = ', v_id_jscet to client;
	--message ' p_numOrder = ', p_numOrder to client;
	call wf_move_invoice_detail (old_server, v_id_jscet, p_numOrder);

	// исправить порядковые 
	// номера позиций для нового и старого счета
	call call_remote(old_server, 'renu_scet', v_id_jscet);
	call call_remote(old_server, 'renu_scet', old_id_jscet);

//	return convert(integer, v_nu_jscet);
	return v_id_jscet;
end;



if exists (select '*' from sysprocedure where proc_name like 'wf_split_jscet') then  
	drop function wf_split_jscet;
end if;

// возвращает id бухгалтерского счета для заказа
// 
create function wf_split_jscet (
	// заказ, который должен быть выделен в отдельный счет
	p_numorder integer
	// номер нового счета
	, p_newInvoice varchar(32) default null
) returns varchar(32)
begin
	set wf_split_jscet = wf_jscet_handle(p_numorder);
	if p_newInvoice is not null then
		update orders set invoice = p_newInvoice where numorder = p_numorder;
	end if;
end;


if exists (select '*' from sysprocedure where proc_name like 'wf_move_jscet') then  
	drop function wf_move_jscet;
end if;

// новый номер бухгалтерского счета для заказа
create function wf_move_jscet (
	// номер заказа - источника, который должен быть перемещен 	
	  p_numorder integer
	// id счета, к которому будет присоединен заказ
	, in p_id_jscet_merge integer
) returns varchar(32)
begin
	set wf_move_jscet = wf_jscet_handle(p_numorder, p_id_jscet_merge);
end;


-------------------------------------
-------------------------------------
-------------------------------------

if exists (select '*' from sysprocedure where proc_name like 'put_jscet') then  
	drop procedure put_jscet;
end if;

create procedure put_jscet (
	  out r_id integer
	, out v_nu_jscet varchar(50)
	, in remoteServerNew varchar(20)
	, in p_numOrder integer
	, in p_id_dest integer
	, in p_nu_old varchar(50) default null 
) 
begin
	declare v_fields varchar(255);
	declare v_values varchar(2000);
	declare r_nu varchar(50);
--	declare v_firm_id integer;
	declare v_invCode varchar(10);
--	declare p_id_dest integer;
	declare v_id_schef integer;
	declare v_id_bux integer;
	declare v_id_bank integer;
	declare v_datev varchar(20);
	declare v_id_cur integer;
	declare v_currency_rate float;
	declare v_order_date varchar(20);
	declare v_check_count integer; 
	declare v_id_jscet integer;
	declare v_intInvoice integer;


	select invCode into v_invCode
	from guideVenture where sysname = remoteServerNew;

/*
	set r_id = select_remote(
		remoteServerNew
		, 'jscet'
		, 'max(id)'
	);
*/

	set v_nu_jscet = nextnu_remote(remoteServerNew, 'jscet', p_nu_old);

	set r_id = r_id + 1;
	set v_order_date = convert(varchar(20), now());
	set v_id_cur = system_currency();
	execute immediate 'call slave_currency_rate_' + remoteServerNew + '(v_datev, v_currency_rate, v_order_date, v_id_cur )';
	
	set v_fields =
		 'nu'
--		+ ', id'
		+ ', rem'
		+ ', id_s'
		+ ', dat' 
		+ ', datv' 
		+ ', state'
		+ ', real_days'
		+ ', id_curr'
		+ ', curr'
//		+ ', id_kad1'
//		+ ', id_kad_bux'
//		+ ', id_s_bank'
		;

	--message 'v_fields  = ', v_fields to client;
	
	set v_values = 
		convert(varchar(20), v_nu_jscet)
--		+ ', ' + convert(varchar(20), r_id)
		+ ', ' + convert(varchar(20), p_numOrder)
		+ ', -1'
		+ ', ''''' + convert(varchar(20), v_order_date, 112) + ''''''
		+ ', ''''' + v_datev + ''''''
		+ ', 1'
		+ ', 3'
		+ ', ' + convert(varchar(20), v_id_cur)
		+ ', ' + convert(varchar(20), v_currency_rate)
		
	;

	--message 'v_values  = ', v_values to client;

	if p_id_dest is not null then
		set v_fields = v_fields
			+ ', id_d'
			+ ', id_d_cargo'
		;
		set v_values = v_values	
			+ ', ' + convert(varchar(20), p_id_dest)
			+ ', ' + convert(varchar(20), p_id_dest)
		;
	end if;

	set r_id = insert_count_remote(remoteServerNew, 'jscet', v_fields, v_values);
end;

-------------------------------------
-------------------------------------
-------------------------------------
if exists (select '*' from sysprocedure where proc_name like 'get_jscet_nu') then  
	drop function get_jscet_nu;
end if;
/*
create function get_jscet_nu (
	remoteServerNew varchar(20)
) returns integer
begin
	declare r_nu varchar(50);
	declare r_id integer;

	set r_id = select_remote(
		remoteServerNew
		, 'jscet'
		, 'max(id)'
	);

	set r_nu = select_remote(
		remoteServerNew
		, 'jscet'
		, 'nu'
		, 'id = ' + convert( varchar(20), r_id)
	);
	set get_jscet_nu = convert(integer, r_nu) + 1;
end;
*/
-------------------------------------------------------------------------
--------------             System      ----------------------------
-------------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_update_System' and tname = 'System') then 
	drop trigger System.wf_update_System;
end if;

create TRIGGER wf_update_System before update on
System
referencing old as old_name new as new_name
for each row
begin
	declare v_fields varchar(1000);
	declare v_values varchar(2000);
	declare v_where varchar(1000);
	declare v_id_cur_rate integer;
	declare v_id_cur integer;
	declare v_currency_rate float;
	declare updated integer;

	if update(kurs) then
		if abs(old_name.kurs) != abs(new_name.kurs) then
			-- update remote bases
			set v_currency_rate = abs(new_name.kurs);
			set v_id_cur_rate = old_name.id_cur_rate;

			set v_fields = 'curse';
			set v_values = 
				'''''' + convert(varchar(20), v_currency_rate) + ''''''
			;
			set v_where = 'id='
				+ convert(varchar(20), v_id_cur_rate) 
				+ ' and dat = ''''' + convert(varchar(20), now(), 112) + '''''';
			;

			set updated = update_count_host(
					'cur_rate'
					, v_fields
					, v_values
					, v_where
			);

			if updated = 0 then
				set v_id_cur_rate = get_nextid('cur_rate');
				set v_fields = 'id, id_cur, dat, curse, rem';
				set v_id_cur = old_name.id_cur;
				set v_values = 
					convert(varchar(20), v_id_cur_rate)
					+', ' + convert(varchar(20), v_id_cur)
					+', ''''' + convert(varchar(20), now(), 112) +''''''
					+', ''''' + convert(varchar(20), v_currency_rate) + ''''''
					+', ''''Установлено в Prior'''''
				;
	
				call insert_host('cur_rate', v_fields, v_values);
				set new_name.id_cur_rate = v_id_cur_rate;

			end if;

		end if;
	end if;

end;



-------------------------------------------------------------------------
--------------             common procs      ----------------------------
-------------------------------------------------------------------------


if exists (select '*' from sysprocedure where proc_name like 'extract_invoice_number') then  
	drop function extract_invoice_number;
end if;

create function extract_invoice_number (
	p_invoice varchar(10)         // номер счета заказа
	,p_invCode varchar(10)        // префикс номера счета для предприятия
) returns varchar(10)
begin
	declare v_invoice varchar(10);
	set v_invoice = substring(p_invoice, 1, char_length(p_invCode));

	if p_invCode is null or char_length(p_invCode) = 0 then
		set extract_invoice_number = p_invoice;
	end if;

//	message 'v_invoice = ', v_invoice to client;

	if p_invCode = v_invoice then 
		set extract_invoice_number = substring(p_invoice, char_length(p_invCode)+1);
	end if;
end;



------------------------------------------------------------------------------------
if exists (select '*' from sysprocedure where proc_name like 'wf_check_jscet_merge') then  
	drop function wf_check_jscet_merge;
end if;

create function wf_check_jscet_merge (
	p_numorder integer            // заказ, которому меняем номер счета
	,p_invoice varchar(10)         // новый номер счета заказа
//	,p_oldInvoice varchar(10)      // прежний номер счета заказа м.б. 'счет ?'
) returns integer
begin
	declare remoteServerOld varchar(32);
	declare varchar_id varchar(20);
	declare v_invoice varchar(10);
	declare f_exists integer;

	// аттрибуты заказа который может быть слит с други
	// (тот у которого руками меняем номер счета)
	declare old_invoice varchar(10);      
	declare old_ventureId integer;
	declare old_id_jscet integer;
	declare old_invCode varchar(20);
	declare old_server varchar(20);
	declare old_firmId integer;


	set wf_check_jscet_merge = 0;

	select invoice, id_jscet, o.ventureId, v.invCode, v.sysname, o.firmId
	into old_invoice, old_id_jscet, old_ventureId, old_invCode, old_server, old_firmId
	from orders o
		join guideventure v on v.ventureId = o.ventureId
	where numorder = p_numorder;

	-- Если есть заказ
	select 0 - count(*) into wf_check_jscet_merge 
		from orders o
		where o.invoice = p_invoice
			and o.numorder != p_numorder
			and isnull(o.shipped, 0) = 0
			and o.ventureId = old_ventureId
			and o.id_jscet is not null and o.id_jscet > 0
			and o.firmId <> old_firmId
		;

	if wf_check_jscet_merge < 0 then
		return;
	end if;

	if old_ventureId is null then
		return;
	end if;

	a:
	for v_server_name as a dynamic scroll cursor for
		select o.numOrder as r_numOrder
			, o.id_jscet as r_id_jscet
		from orders o
		where o.invoice = p_invoice
			and o.numorder != p_numorder
			and isnull(o.shipped, 0) = 0
			and o.ventureId = old_ventureId
			and o.id_jscet is not null and o.id_jscet > 0
	do

		set wf_check_jscet_merge = r_id_jscet;
		leave a;
/*
		set v_invoice = extract_invoice_number(v_invoice, old_invCode);

		set varchar_id = select_remote(old_server, 'jscet', 'max(id)', 'nu = ''''' + v_invoice + '''''');
		set wf_check_jscet_merge = convert(integer, varchar_id);
		if r_id_jscet != wf_check_jscet_merge then
			// есть такой заказ, у которого id счета другой
			// а номер такой же, на который мы хотим перевести заказ p_numOrder
			// Ситуация для слияния заказа в один
			set f_exists = 1;
		else
			// ни о чем не говорит. Это нормальная ситуация, 
			// к примеру, сливается третий заказ в один счет
		end if;
*/
	end for;

end;


-------------------------------------
-------------------------------------
-------------------------------------
if exists (select '*' from sysprocedure where proc_name like 'wf_merge_jscet') then  
	drop procedure wf_merge_jscet;
end if;

create procedure wf_merge_jscet (
	  p_numorder integer			// заказ, которому меняем номер счета
	, p_id_jscet_new integer        // id счета бухгалтерской базы
	, p_nu_jscet varchar(32)        // номер бух. счета
)
begin
	declare v_updated integer;
	// аттрибуты заказа который может быть слит с други
	// (тот у которого руками меняем номер счета)
	declare old_invoice varchar(10);      
	declare old_ventureId integer;
	declare old_id_jscet integer;
	declare old_invCode varchar(20);
	declare old_server varchar(20);
	declare scet_nu varchar(20);
	declare v_blank_inv integer;

	
	select invoice, id_jscet, o.ventureId, v.invCode, v.sysname 
	into old_invoice, old_id_jscet, old_ventureId, old_invCode, old_server
	from orders o
		join guideventure v on v.ventureId = o.ventureId
	where numorder = p_numorder;

	if old_ventureId is null then
		return;
	end if;
	set scet_nu = select_remote(
		old_server
		, 'scet'
		, 'max(nu)'
		, 'id_jmat = ' + convert(varchar(20), p_id_jscet_new)
	);
	if scet_nu is not null then
		call update_remote(old_server, 'scet', 'nu'
			, 'nu + ' + convert(varchar(20), scet_nu)
			, 'id_jmat = ' + convert (varchar(20), old_id_jscet)
		);
	end if;

	set v_updated = update_count_remote(old_server, 'scet', 'id_jmat'
		, convert(varchar(20), p_id_jscet_new)
		, 'id_jmat = ' + convert (varchar(20), old_id_jscet)
	);

	--message 'old_id_jscet = ', old_id_jscet to client;

	select count(*) into v_blank_inv from orders where id_jscet = old_id_jscet;

	--message 'v_blank_inv = ', v_blank_inv to client;
	if v_blank_inv = 1 then
		-- Исправление баги: счет не обязательно удалять,
		-- а только если не осталось на него ни одной ссылки
		call delete_remote( old_server, 'jscet', 'id = ' + convert(varchar(20), old_id_jscet));
	end if;

	update orders set id_jscet = p_id_jscet_new where numorder = p_numorder;
	update orders set invoice = p_nu_jscet where numorder = p_numorder;

end;


-------------------------------------
-------------------------------------
-------------------------------------
if exists (select '*' from sysprocedure where proc_name like 'gualify_guide') then  
	drop procedure gualify_guide;
end if;

create procedure gualify_guide (
	  p_id_guide_jmat integer
	, out p_tp1 integer
	, out p_tp2 integer
	, out p_tp3 integer
	, out p_tp4 integer
) 
begin
		if p_id_guide_jmat = 1127 then 
		-- приход валютный
			set p_tp1 = 1; set p_tp2 = 1; set p_tp3 = 2; set p_tp4 = 7; 
		elseif p_id_guide_jmat = 1120 then 
		-- приход рублевый
			set p_tp1 = 1; set p_tp2 = 1; set p_tp3 = 2; set p_tp4 = 0;
		elseif p_id_guide_jmat = 1220 then 
		-- межсклад
			set p_tp1 = 2; set p_tp2 = 2; set p_tp3 = 2; set p_tp4 = 0;
		elseif p_id_guide_jmat = 1210 then 
		-- расход
			set p_tp1 = 3; set p_tp2 = 2; set p_tp3 = 1; set p_tp4 = 0; 
		elseif p_id_guide_jmat = 1023 then 
		-- инвентаризация
			set p_tp1 = 0; set p_tp2 = 0; set p_tp3 = 2; set p_tp4 = 3; 
		end if;
end;



if exists (select '*' from sysprocedure where proc_name like 'wf_get_comtex_tp') then  
	drop function wf_get_comtex_tp;
end if;

create function wf_get_comtex_tp (
	p_id_guide_jmat integer
) returns varchar(20)
begin
	declare v_ret_char varchar(20);
	-- приход
	if p_id_guide_jmat = 1120 then
		return '1,1,2,0';
	end if;
	-- межсклад
	if p_id_guide_jmat = 1220 then
		return '2,2,2,0';
	end if;
	-- расход
	if p_id_guide_jmat = 1210 then
		return '3,2,1,0';
	end if;
	-- инвентаризация
	if p_id_guide_jmat = 1023 then
		return '0,0,2,3';
	end if;
	set v_ret_char = convert(varchar(20), p_id_guide_jmat);
	return		substring (v_ret_char, 1, 1)
		+ ', '+ substring (v_ret_char, 2, 1) 
		+ ', '+ substring (v_ret_char, 3, 1) 
		+ ', '+ substring (v_ret_char, 4, 1) 
	;

end;



--------------------------
if exists (select '*' from sysprocedure where proc_name like 'wf_insert_jmat') then  
	drop procedure wf_insert_jmat;
end if;

create procedure wf_insert_jmat (
		p_srvName varchar(20)
		, p_id_guide_jmat integer
		, p_id_jmat integer
		, p_jmat_date date
		, p_jmat_nu integer
		, p_osn varchar(100)
		, p_id_currency integer
		, p_datev date
		, p_currency_rate float
		, p_id_s integer
		, p_id_d integer
		, p_id_jscet integer default 0
		, p_id_code integer default 0
)
begin
	declare v_fields varchar(255);
	declare v_values varchar(2000);
	declare v_tp varchar(20);
--	declare out_id integer;
	set v_tp = wf_get_comtex_tp(p_id_guide_jmat);
	set p_id_jscet = isnull(p_id_jscet, 0);
	set p_id_code = isnull(p_id_code, 0);


	set v_fields = 'id'
		+ ', dat'
		+ ' , nu '
		+ ', id_s'
		+ ', id_d'
		+ ', osn'
		+ ', id_guide'
		+ ', tp1, tp2, tp3, tp4'
		+ ', id_curr'
		+ ', datv'
		+ ', curr'
		+ ', id_jscet'
		+ ', id_code'

	;   
	set v_values = convert(varchar(20), p_id_jmat)
		+ ', ''''' + convert(varchar(20), p_jmat_date) + ''''''
		+ ', ' + convert(varchar(20), p_jmat_nu)
		+ ', ' + convert(varchar(20), p_id_s)
		+ ', ' + convert(varchar(20), p_id_d)
		+ ', ''''' + p_osn + ''''''
		+ ', ' + convert(varchar(20), p_id_guide_jmat)
		+ ', ' + v_tp
		+ ', ' + convert(varchar(20), p_id_currency)
		+ ', ''''' + convert(varchar(20), p_datev, 112) + ''''''
		+ ', ' + convert(varchar(20), p_currency_rate)
		+ ', ' + convert(varchar(20), p_id_jscet)
		+ ', ' + convert(varchar(20), p_id_code)

	;
	message 'Накладная JMAT:fields = ', v_fields to client;
	message '	values = ', v_values to client;
	execute immediate 'call slave_insert_'+ p_srvName +' (''jmat'', ''' +v_fields + ''', ''' + v_values + ''')'

end;




if exists (select '*' from sysprocedure where proc_name like 'wf_insert_mat') then  
	drop function wf_insert_mat;
end if;

create function wf_insert_mat (
		p_srvName varchar(20)
		, p_id_mat integer
		, p_id_jmat integer
		, p_id_inv integer
		, p_mat_nu integer
		, p_quant float
		, p_cena float
		, p_currency_rate float
		, p_id_s integer
		, p_id_d integer
		, p_perList float default 1
--		, p_cenav float
--		, p_date date
--		, p_id_cur integer
--		, p_datev varchar(20)

) returns integer
begin
	declare v_fields varchar(255);
	declare v_values varchar(2000);
	declare v_tp varchar(20);
	declare v_tp1 integer;
	declare v_tp2 integer;
	declare v_tp3 integer;
	declare v_tp4 integer;
	declare v_id_guide integer;

	if p_id_mat is null then
		call slave_nextid_stime('mat', p_id_mat);
	end if;

	set wf_insert_mat = p_id_mat;

	call slave_select_stime(v_id_guide, 'jmat', 'id_guide', 'id = ' + convert(varchar(20), p_id_jmat));
	call slave_select_stime(v_tp1, 'jmat', 'tp1', 'id = ' + convert(varchar(20), p_id_jmat));
	call slave_select_stime(v_tp2, 'jmat', 'tp2', 'id = ' + convert(varchar(20), p_id_jmat));
	call slave_select_stime(v_tp3, 'jmat', 'tp3', 'id = ' + convert(varchar(20), p_id_jmat));
	call slave_select_stime(v_tp4, 'jmat', 'tp4', 'id = ' + convert(varchar(20), p_id_jmat));

//	set v_tp = wf_get_comtex_tp(v_id_guide);
	set v_tp = convert(varchar(20), v_tp1)
		+ ',' + convert(varchar(20), v_tp2)
		+ ',' + convert(varchar(20), v_tp3)
		+ ',' + convert(varchar(20), v_tp4)
	;


	set v_fields = 'id'
		+ ', id_jmat'
		+ ', id_inv'
		+ ', nu'
		+ ', id_s'
		+ ', id_d'
		+ ', kol1'
		+ ', kol3'
		+ ', kol2'
		+ ', kol23'
		+ ', tp1, tp2, tp3, tp4'
		+ ', summa'
		+ ', summa_sale'
	;
	if v_id_guide = 1127 then
	--  "приход по импорту в валюте"
		set v_fields = v_fields
			+ ', summav'
			+ ', summa_salev'
		;
	end if;

	set v_values = convert(varchar(20), p_id_mat)
		+ ', ' + convert(varchar(20), p_id_jmat)
		+ ', ' + convert(varchar(20), p_id_inv)
		+ ', ' + convert(varchar(20), p_mat_nu)
		+ ', ' + convert(varchar(20), p_id_s)
		+ ', ' + convert(varchar(20), p_id_d)
		+ ', ' + convert(varchar(20), p_quant / p_perList)
		+ ', ' + convert(varchar(20), p_quant / p_perList)
		+ ', ' + convert(varchar(20), p_quant / p_perList)
		+ ', ' + convert(varchar(20), p_quant / p_perList)
		+ ', ' + v_tp
		+ ', ' + convert(varchar(20), p_quant* p_cena * p_currency_rate / p_perList)
		+ ', ' + convert(varchar(20), p_quant* p_cena * p_currency_rate / p_perList)
	;

	if v_id_guide = 1127 then
	-- приход по импорту в валюте 
		set v_values = v_values 
			+ ', ' + convert(varchar(20), p_quant * p_cena / p_perList)
			+ ', ' + convert(varchar(20), p_quant * p_cena / p_perList)
		;
	end if;
--	message 'Предметы накладной:fields = ', v_fields to client;
--	message '	values = ', v_values to client;
	call insert_remote(p_srvName, 'mat', v_fields, v_values);
--	execute immediate 'call slave_insert_'+ p_srvName +' (''mat'', ''' +v_fields + ''', ''' + v_values + ''')'

end;


if exists (select '*' from sysprocedure where proc_name like 'wf_insert_scet') then  
	drop procedure wf_insert_scet;
end if;

create function wf_insert_scet (
		  p_srvName varchar(20)
		, p_id_jscet integer
		, p_id_inv integer
		, p_quant float
		, p_cena float
		, p_date date
)
returns integer
begin
	declare v_id_scet integer;
	declare v_fields varchar(255);
	declare v_values varchar(2000);
	declare scet_nu integer;
	declare v_currency_rate float;
	declare v_datev varchar(20);
	declare v_id_cur integer;


--	set p_quant = round(p_quant, 2);
--	set p_cena = round(p_cena, 2);

 
  if p_srvName is not null and p_id_jscet is not null then
//	execute immediate 'select max(nu)+1 into scet_nu from scet_' + p_srvName + ' where id_jmat = ' + convert(varchar(20), p_id_jscet);

	-- Получить следующий порядковый номер счета бух.базы
	set scet_nu = select_remote(
		p_srvName
		, 'scet'
		, 'max(nu)+1'
		, 'id_jmat = ' + convert(varchar(20), p_id_jscet)
	);

	set scet_nu = isnull(scet_nu, 1);

	-- По какому курсу, учитывая, что в бухгалтерии только рубли, а в Приоре - УЕ
	set v_id_cur = system_currency();

	execute immediate 'call slave_currency_rate_' + p_srvName + '(v_datev, v_currency_rate, p_date, v_id_cur )';
	
	set v_fields = '
		 id_jmat
		,id_inv
		,kol1
		,nu
		,summa_sale
		,summa_salev
	';

	set v_values = 
		convert(varchar(20), p_id_jscet)
		+', '+ convert(varchar(20), p_id_inv)
		+', '+ convert(varchar(20), p_quant)
		+', '+ convert(varchar(20), scet_nu)
		+', '+ convert(varchar(20), round(p_quant*p_cena*v_currency_rate, 2))
		+', '+ convert(varchar(20), round(p_quant*p_cena, 2))
	;
	--message 'p_cena = ', p_cena to client;
	--message 'p_quant = ', p_quant to client;
	--message 'v_values = ', v_values to client;

	-- изменения в бухгалтерской базе данных
	set v_id_scet = insert_count_remote(p_srvName, 'scet', v_fields, v_values);

//	execute immediate 'select id into v_id_scet from scet_' + p_srvName + ' s where s.id_jmat = p_id_jscet and s.id_inv = p_id_inv';
/*
	set v_id_scet = select_remote(
		p_srvName
		, 'scet s'
		, 'id'
		, 's.id_jmat = '+convert(varchar(20), p_id_jscet) + ' and s.id_inv = ' +convert(varchar(20), p_id_inv)
	);
*/
	return v_id_scet;
  end if;
  return null;
	
end;





/************************************************************/
/*                 HOST PROCEDURES                          */
/************************************************************/



if exists (select '*' from sysprocedure where proc_name like 'get_nextid') then
	drop function get_nextid;
end if;

create function get_nextid(table_name varchar(100)) returns integer
/*
	получает следующий свободный id для таблицы table_name с учетом всех
*/
begin
	declare curId integer;
	declare maxId integer;
	set maxId = 0;set curId = 0;
	
  for v_server_name as a dynamic scroll cursor for
	select srvname as cur_server from sys.sysservers s join guideventure v on s.srvname = v.sysname and v.standalone = 0 do
	
	execute immediate 'call slave_nextid_' + cur_server + '('''+table_name+''', curId)';
	if maxId < curId then
		set maxId = curId;
	end if;
  end for;
  return maxId;
end;

/************************************************************/
/*                  PRIOR SPECIFIC PROCS                    */
/************************************************************/


if exists (select '*' from sysprocedure where proc_name like 'wf_set_invoice_detail') then  
	drop procedure wf_set_invoice_detail;
end if;


-- Процедура синхронизирует предметы заказа Приора
-- с предметами счета в бухгалтерской базе комтеха
-- Это нужно сделать, если в заказ сначала 
-- добавть предметы, а только потом назначить предприятие,
-- через которую этот заказ должен пройти.
create procedure wf_set_invoice_detail (
			p_srvName varchar(20)
			, p_id_jscet integer
			, p_numOrder integer
			, p_date date
)
begin

	declare v_id_scet integer;
	declare v_id_inv integer;
	declare is_variant integer;
	declare v_id_variant integer;
	declare is_uslug integer;
	declare v_quant float;
	declare v_perList float;

	set is_uslug = 1; // предполагаем изначально, что да


	for c_nomenk as n dynamic scroll cursor for
		select 
			  p.nomNom as r_nomNom
			, p.quant as r_quant
			, p.cenaEd as r_cenaEd
		from xPredmetybynomenk p
		where p.numOrder = p_numOrder
	do
	    set is_uslug = 0; -- есть предметы к заказу, значит не услуга

		select id_inv, perList into v_id_inv, v_perList from sGuideNomenk where nomnom = r_nomNom;
		
		set v_id_scet = 
			wf_insert_scet(
				p_srvName
				, p_id_jscet
				, v_id_inv
				, r_quant / v_perList
				, r_cenaEd
				, p_date
			);
		update xPredmetyByNomenk set id_scet = v_id_scet where current of n;

	end for;


	for c_izd as i dynamic scroll cursor for
		select 
			  prId as r_prId
			, prExt as r_prExt
			, quant as r_quant
			, cenaEd as r_cenaEd
		from xPredmetyByIzdelia p
		where p.numOrder = p_numOrder
	do

	    set is_uslug = 0; -- есть предметы к заказу, значит не услуга
		select id_inv into v_id_inv from sGuideProducts where prId = r_prId;

		-- смотрим, является ли изделие вариантным?
		
		select count(*) into is_variant from sVariantPower where productId = r_prId;
		if is_variant = 1 then
			-- ищем и/или добавляем вариант в Inv
			set v_id_variant = wf_get_variant_id(p_numOrder, r_prId, r_prExt);
			select id_inv into v_id_inv 
			from sGuideComplect 
			where 
				id_variant = v_id_variant;
		end if;

		set v_id_scet = 
			wf_insert_scet(
				p_srvName
				, p_id_jscet
				, v_id_inv
				, r_quant
				, r_cenaEd
				, p_date
			);

		update xPredmetyByIzdelia set id_scet = v_id_scet, id_inv = v_id_inv where current of i;
	end for;  -- цикла по изделиям

	select ordered into v_quant from orders where numorder = p_numOrder;
	if is_uslug = 1 and abs(v_quant) > 0.001 then
		-- ищем товар под названием "услуга"
		select id_inv into v_id_inv from sGuideNomenk where nomNom = 'УСЛ';


		set v_id_scet = 
			wf_insert_scet(
				p_srvName
				, p_id_jscet
				, v_id_inv
				, 1 // quant
				, v_quant//r_cenaEd
				, now()//p_date
			);

	end if;


end;




if exists (select '*' from sysprocedure where proc_name like 'wf_move_invoice_detail') then  
	drop procedure wf_move_invoice_detail;
end if;


-- Процедура написана на основе wf_move_invoice_detail (через Copy&Paste)
-- 
-- Только вместо добаления предметов перепривязываем позицию к другому счету
create procedure wf_move_invoice_detail (
	p_srvName varchar(20)
	, p_id_jscet_new integer
	, p_numOrder integer
)
begin

	declare is_uslug integer;
	declare v_updated integer;
	declare v_quant float;
	declare v_id_scet integer;
	declare v_id_jscet integer;
	declare v_id_inv integer;

	set is_uslug = 1; // предполагаем изначально, что да

	for c_nomenk as n dynamic scroll cursor for
		select 
			id_scet as r_id_scet
		from xPredmetybynomenk p
		where p.numOrder = p_numOrder
	do
	    set is_uslug = 0; -- есть предметы к заказу, значит не услуга

		set v_updated = update_count_remote(p_srvName, 'scet', 'id_jmat'
			, convert(varchar(20), p_id_jscet_new)
			, 'id = ' + convert (varchar(20), r_id_scet)
		);


	end for;


	for c_izd as i dynamic scroll cursor for
		select 
			id_scet as r_id_scet
		from xPredmetyByIzdelia p
		where p.numOrder = p_numOrder
	do

	    set is_uslug = 0; -- есть предметы к заказу, значит не услуга

		set v_updated = update_count_remote(p_srvName, 'scet', 'id_jmat'
			, convert(varchar(20), p_id_jscet_new)
			, 'id = ' + convert (varchar(20), r_id_scet)
		);

	end for;  -- цикла по изделиям

	--message 'is_uslug = ', is_uslug to client;
	select ordered into v_quant from orders where numorder = p_numOrder;
	if is_uslug = 1 then
		-- Искать услугу ровно с такой же суммой
		-- относящуюся к старому счету и перепривязываем ее к новому

		select id_jscet into v_id_jscet from orders where numorder = p_numorder;

		-- ищем товар под названием "услуга"
		select id_inv into v_id_inv from sGuideNomenk where nomNom = 'УСЛ';

		--message 'v_id_jscet     = ', v_id_jscet    to client;
		--message 'p_id_jscet_new = ', p_id_jscet_new to client;
		--message 'v_quant        = ', v_quant       to client;
		--message 'v_id_inv       = ', v_id_inv      to client;

		call call_remote(p_srvName, 'move_uslug', 
			         convert(varchar(20), v_id_jscet    )
			+ ', ' + convert(varchar(20), p_id_jscet_new)
			+ ', ' + convert(varchar(20), isnull(v_quant, 0)       )
			+ ', ' + convert(varchar(20), v_id_inv      )
		);

/*
		set v_id_scet = select_remote(
			p_srvName
			, 'scet'
			, 'id'
			, 'id_jmat = '+ convert(varchar(20), p_id_jscet_new)
				+ ' and id_inv = ' + convert(varchar(20), v_id_inv)
				+ ' and summa_salev = ' + convert(varchar(20), v_quant)
		);

		if v_id_scet is not null then
			set v_updated = update_count_remote(p_srvName, 'scet', 'id_jmat'
				, convert(varchar(20), p_id_jscet_new)
				, 'id = ' + convert (varchar(20), v_id_scet)
			);
		end if;


		set v_id_scet = 
			wf_insert_scet(
				p_srvName
				, p_id_jscet
				, v_id_inv
				, 1 // quant
				, v_quant//r_cenaEd
				, now()//p_date
			);
*/
	end if;


end;





-- Получить ид единицы измерения. ид является общим на все базы
-- Если такой единицы еще нет, то она добавляется во все базы
if exists (select '*' from sysprocedure where proc_name like 'wf_getEdizmId') then  
	drop procedure wf_getEdizmId;
end if;

create FUNCTION wf_getEdizmId (edizm varchar(100), p_rem varchar(100) default 'created by stime') returns integer
begin
	declare edizmId integer;
	declare v_values varchar(200);
	select id_edizm into edizmId from edizm where name = edizm;
	if edizmId is not null then
		return edizmId;
	end if;

	set edizmId = get_nextId('edizm');
	set v_values = convert(varchar(20), edizmId) 
		+ ', ''''' + edizm + ''''''
		+ ', '''''+p_rem+'''''';

	call insert_host('edizm', 'id, nm,rem', v_values );
	insert into edizm (id_edizm, name) 
	values (edizmId, edizm);
	
	return edizmId;
end;


-- Получить ид размера. ид является общим на все базы
-- Если такога размера еще нет, то создается новый размер
-- и добавляется во все базы
if exists (select '*' from sysprocedure where proc_name like 'wf_getSizeId') then  
	drop procedure wf_getSizeId;
end if;

create FUNCTION wf_getSizeId (sz varchar(100), p_rem varchar(100) default 'created by stime') returns integer
begin
	declare sizeId integer;
	declare v_values varchar(200);

	select id_size into sizeId from size where name = sz;
	if sizeId is not null then
		return sizeId;
	end if;

	set sizeId = get_nextId('size');
	set v_values = convert(varchar(20), sizeId)
		+ ', ''''' + sz + ''''''
		+ ', '''''+p_rem+'''''';

	call insert_host('size', 'id,nm,rem', v_values );
	insert into size (id_size, name)
	values (sizeId, sz);
	return sizeId;
end;




-------------------------------------------------------------------------
--------------             xPredmetyByIzdelia      ----------------------
-------------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_insert_izd' and tname = 'xPredmetyByIzdelia') then 
	drop trigger xPredmetyByIzdelia.wf_insert_izd;
end if;

create TRIGGER wf_insert_izd before insert on
xPredmetyByIzdelia
referencing new as new_name
for each row
begin
	declare v_id_scet integer;
	declare v_id_jscet integer;
	declare v_id_inv integer;
--	declare v_ventureid integer;
	declare remoteServerNew varchar(32);
	declare v_invCode varchar(10);
	declare v_fields varchar(255);
	declare v_values varchar(2000);
	declare v_date date;
 
	select id_jscet, inDate, sysname, invCode 
	into v_id_jscet, v_date, remoteServerNew, v_invcode  
		from orders o
		join GuideVenture v on o.ventureId = v.ventureId and v.standalone = 0
		where numOrder = new_name.numOrder;

	select id_inv into v_id_inv 
		from sGuideProducts where prId = new_name.prId;
  
	if remoteServerNew is not null and v_id_jscet is not null then
		set v_id_scet =	
			wf_insert_scet (
				remoteServerNew
				, v_id_jscet
				, v_id_inv
				, new_name.quant
				, new_name.cenaEd
				, v_date
			);
		set new_name.id_scet = v_id_scet;
		set new_name.id_inv = v_id_inv;
	end if;
end;



if exists (select 1 from systriggers where trigname = 'wf_update_izd' and tname = 'xPredmetyByIzdelia') then 
	drop trigger xPredmetyByIzdelia.wf_update_izd;
end if;

create TRIGGER wf_update_izd before update on
xPredmetyByIzdelia
referencing old as old_name new as new_name
for each row
begin
	declare v_id_scet integer;
	declare v_numorder integer;
	declare v_belong_id integer;
	declare remoteServerNew varchar(32);
	declare v_values varchar(100);
	declare v_fields varchar(200);
	declare v_currency_rate float;
	
	set v_id_scet = old_name.id_scet;
--	set v_numorder = old_name.numOrder;

	select sysname
	into remoteServerNew
		from orders o
		join GuideVenture v on o.ventureId = v.ventureId and v.standalone = 0
		where numOrder = old_name.numOrder;


	if remoteServerNew is not null then
		if update(quant) or update(cenaEd) then
			set v_currency_rate = system_currency_rate();
			call update_remote(remoteServerNew, 'scet', 'summa_sale'
				, convert(varchar(20), v_currency_rate * new_name.quant*new_name.cenaEd)
				, 'id = ' + convert(varchar(20), v_id_scet)
			);
			call update_remote(remoteServerNew, 'scet', 'summa_salev'
				, convert(varchar(20), new_name.quant*new_name.cenaEd)
				, 'id = ' + convert(varchar(20), v_id_scet)
			);
//			call update_remote(remoteServerNew, 'scet', 'summa_sale', convert(varchar(20), new_name.quant*new_name.cenaEd), 'id = ' + convert(varchar(20), v_id_scet));
		end if;
		if update(quant) then
			call update_remote(remoteServerNew, 'scet', 'kol1', convert(varchar(20), new_name.quant), 'id = ' + convert(varchar(20), v_id_scet));
//			call update_remote(remoteServerNew, 'scet', 'kol3', convert(varchar(20), new_name.quant), 'id = ' + convert(varchar(20), v_id_scet));
		end if;
	end if;
  
end;


if exists (select 1 from systriggers where trigname = 'wf_delete_izd' and tname = 'xPredmetyByIzdelia') then 
	drop trigger xPredmetyByIzdelia.wf_delete_izd;
end if;

create TRIGGER wf_delete_izd before delete on
xPredmetyByIzdelia
referencing old as old_name
for each row
begin
	declare remoteServerNew varchar(32);
	select sysname
	into remoteServerNew
		from orders o
		join GuideVenture v on o.ventureId = v.ventureId and v.standalone = 0
		where numOrder = old_name.numOrder;

	if remoteServerNew is not null then
		call delete_remote(remoteServerNew, 'scet', 'id = ' + convert(varchar(20), old_name.id_scet));
	end if;
end;



-------------------------------------------------------------------------
--------------             xPredmetyByNomenk      -----------------------
-------------------------------------------------------------------------
--select * from scet_pm order by id_jmat desc
--select * from xpredmetybynomenk order by 1 desc
--select max(nu)+1  from scet_pm where id_jmat = 13281



if exists (select 1 from systriggers where trigname = 'wf_insert_nomenk' and tname = 'xPredmetyByNomenk') then 
	drop trigger xPredmetyByNomenk.wf_insert_nomenk;
end if;

create TRIGGER wf_insert_nomenk before insert on
xPredmetyByNomenk
referencing new as new_name
for each row
begin
	declare v_id_scet integer;
	declare v_id_jscet integer;
	declare v_id_inv integer;
	declare v_ventureid integer;
	declare remoteServerNew varchar(32);
	declare v_invCode varchar(10);
	declare v_fields varchar(255);
	declare v_values varchar(2000);
	declare scet_nu integer;
	declare v_date date;
	declare v_perList float;

	select id_jscet, ventureId, inDate  into v_id_jscet, v_ventureId, v_date from orders where numOrder = new_name.numOrder;
	select id_inv, perList into v_id_inv, v_perList from sGuideNomenk where nomNom = new_name.nomNom;
	select sysname, invCode into remoteServerNew, v_invcode from GuideVenture where ventureId = v_ventureId;

	if remoteServerNew is not null and v_id_jscet is not null then
	  -- Заказ, который имеет ссылки в бух.базах интеграции
	  -- т.е. уже назначен той, иди другой фирме
		set new_name.id_scet = 
			wf_insert_scet(
				remoteServerNew
				, v_id_jscet
				, v_id_inv
				, new_name.quant / v_perList
				, new_name.cenaEd
				, v_date
			);
	end if;
	  
end;


if exists (select 1 from systriggers where trigname = 'wf_update_nomenk' and tname = 'xPredmetyByNomenk') then 
	drop trigger xPredmetyByNomenk.wf_update_nomenk;
end if;

create TRIGGER "wf_update_nomenk" before update on
xPredmetyByNomenk
referencing old as old_name new as new_name
for each row
begin
	declare v_id_scet integer;
	declare v_belong_id integer;
	declare remoteServerNew varchar(32);
	declare v_currency_rate float;

	declare v_values varchar(100);
	declare v_fields varchar(200);
	
	set v_id_scet = old_name.id_scet;

	select sysname
	into remoteServerNew
		from orders o
		join GuideVenture v on o.ventureId = v.ventureId and v.standalone = 0
		where numOrder = old_name.numOrder;


	if remoteServerNew is not null then
		if update(quant) or update(cenaEd) then
			set v_currency_rate = system_currency_rate();
			call update_remote(remoteServerNew, 'scet', 'summa_sale'
				, convert(varchar(20), v_currency_rate * new_name.quant*new_name.cenaEd)
				, 'id = ' + convert(varchar(20), v_id_scet)
			);
			call update_remote(remoteServerNew, 'scet', 'summa_salev'
				, convert(varchar(20), new_name.quant*new_name.cenaEd)
				, 'id = ' + convert(varchar(20), v_id_scet)
			);
        end if;
		if update(quant) then
			call update_remote(remoteServerNew, 'scet', 'kol1', convert(varchar(20), new_name.quant), 'id = ' + convert(varchar(20), v_id_scet));
			call update_remote(remoteServerNew, 'scet', 'kol3', convert(varchar(20), new_name.quant), 'id = ' + convert(varchar(20), v_id_scet));
		end if;
	end if;
	  
end;
	
	
if exists (select 1 from systriggers where trigname = 'wf_delete_nomenk' and tname = 'xPredmetyByNomenk') then 
	drop trigger xPredmetyByNomenk.wf_delete_nomenk;
end if;
    
create TRIGGER "wf_delete_nomenk" before delete on
xPredmetyByNomenk
referencing old as old_name
for each row
begin
	declare remoteServerNew varchar(32);
	select sysname
	into remoteServerNew
		from orders o
		join GuideVenture v on o.ventureId = v.ventureId and v.standalone = 0
		where numOrder = old_name.numOrder;

	if remoteServerNew is not null then
		call delete_remote(remoteServerNew, 'scet', 'id = ' + convert(varchar(20), old_name.id_scet));
	end if;
end;



-------------------------------------------------------------------------
--------------             xVariantNomenc      --------------------------
-------------------------------------------------------------------------

if exists (select 1 from systriggers where trigname = 'wf_insert_variant' and tname = 'xVariantNomenc') then 
	drop trigger xVariantNomenc.wf_insert_variant;
end if;

create TRIGGER wf_insert_variant before insert on
xVariantNomenc
referencing new as new_name
for each row
begin
	declare v_id_scet integer;
	declare v_id_jscet integer;
	declare v_id_inv integer;
	declare v_ventureid integer;
	declare remoteServerNew varchar(32);
	declare v_invCode varchar(10);
--	declare v_fields varchar(255);
--	declare v_values varchar(2000);
	declare curNo integer;
	declare v_quant float;
	declare v_cenaEd float;
	declare v_total integer;
	declare v_id_variant integer;
 
    -- Сколько строчек уже вставлено?
	select count(*) into curNo 
	from xVariantNomenc 
	where
		numOrder = new_name.numOrder 
		and prid = new_name.prid 
		and prExt = new_name.prExt;

	-- А сколько нужно"?"
	select numgroup into v_total from sVariantPower where productId = new_name.prid;

	-- поскольку триггер не after, а before, сумма должна быть на единицу меньше
	if curNo + 1 != v_total then
		--еще не все строки вариантного изделия добавлены
		-- ждем, когда будут добавлены все!
		return;
	end if;

	-- Ищем (и добавляем автоматом) реализацию варианта
	set v_id_variant= wf_get_variant_Id(
			 new_name.numOrder
			,new_name.prId
			,new_name.prExt
			,new_name.nomNom
		);

	select id_inv into v_id_inv 
	from sguidecomplect 
	where id_variant = v_id_variant;
	
	select 
		quant
		, cenaEd 
		, id_scet
	into v_quant
		, v_cenaEd 
		, v_id_scet
	from xPredmetyByIzdelia i 
	where
		i.numOrder = new_name.numOrder 
		and i.prId = new_name.prId 
		and i.prExt = new_name.prExt
	;


	select id_jscet, ventureId  into v_id_jscet, v_ventureId from orders where numOrder = new_name.numOrder;
--	select id_inv into v_id_inv from sGuideProducts where prId = new_name.prId;
	select sysname, invCode into remoteServerNew, v_invcode from GuideVenture where ventureId = v_ventureId;
  
	if remoteServerNew is not null and v_id_scet is not null then
	-- Заказ, который имеет ссылки в бух.базах интеграции
	-- т.е. уже назначен той, иди другой фирме
		call update_remote(remoteServerNew, 'scet', 'id_inv', convert(varchar(20), v_id_inv), 'id = ' + convert(varchar(20), v_id_scet));
--		update xPredmetyByIzdelia i set id_scet = v_id_scet where
--			i.numOrder = new_name.numOrder and i.prId = new_name.prId and i.prExt = new_name.prExt;
	end if;
	
end;


-------------------------------------------------------------------------
--------------             sGuideKlass      ----------------------------
-------------------------------------------------------------------------

if exists (select 1 from systriggers where trigname = 'wf_insert_klass' and tname = 'sguideklass') then 
	drop trigger sguideklass.wf_insert_klass;
end if;

create TRIGGER "wf_insert_klass" before insert on
sguideklass
referencing new as new_name
for each row
begin
	declare v_id_inv integer;
	declare v_values varchar(200);
	declare v_belong_id integer;
	set v_id_inv = get_nextid('inv');
	select id_inv into v_belong_id from sguideklass where klassId = new_name.parentKlassId;

	set v_values = convert(varchar(20), v_id_inv)
		+ ', ' + convert(varchar(20), v_belong_id)
		+ ', ''''' + new_name.klassname + ''''''
		+ ', 1'
	;
	
	call insert_host('inv', 'id, belong_id, nm, is_group', v_values);
	set new_name.id_inv=v_id_inv;
/*
	insert into inv (klassid, parentklassid, NM, is_group)
	select 
		new_name.klassid
		, new_name.parentklassid
		, new_name.klassname
		, 1;
*/
end;



if exists (select 1 from systriggers where trigname = 'wf_update_klass' and tname = 'sguideklass') then 
	drop trigger sguideklass.wf_update_klass;
end if;

create TRIGGER "wf_update_klass" before update on
sguideklass
referencing old as old_name new as new_name
for each row
begin
	declare v_id_inv integer;
	declare v_belong_id integer;

	declare v_values varchar(100);
	declare v_fields varchar(200);
	
	set v_id_inv = old_name.id_inv;
	
  if update(klassname) then
	call update_host('inv', 'nm', '''''' + new_name.klassName + '''''', 'id = ' + convert(varchar(20), v_id_inv));
--    update inv as pi set
--      nm = new_name.klassName where
--      pi.id = old_name.id_inv
  end if;
  if update(parentklassId) then
	select id_inv into v_belong_id from sguideklass where klassid = new_name.parentklassId;
	call update_host('inv', 'belong_id', convert(varchar(20), v_belong_id), 'id = ' + convert(varchar(20), v_id_inv));
--    update inv as pi set
--      belong_id = p.id_Inv
--	from sguideklass p
--	where
--      pi.id = old_name.id_inv
--	and p.klassId = new_name.parentklassId
  end if;
  
end;


if exists (select 1 from systriggers where trigname = 'wf_delete_klass' and tname = 'sGuideKlass') then 
	drop trigger sGuideKlass.wf_delete_klass;
end if;

create TRIGGER "wf_delete_klass" before delete on
sGuideKlass
referencing old as old_name
for each row
begin
	if old_name.id_inv is not null then
		call delete_host('inv', 'id = ' + convert(varchar(20), old_name.id_inv));
	end if;
--  delete from inv where id = old_name.id_inv;
end;



-------------------------------------------------------------------------
--------------             sGuideSeries      ----------------------------
-------------------------------------------------------------------------


if exists (select 1 from systriggers where trigname = 'wf_insert_seria' and tname = 'sguideseries') then 
	drop trigger sguideseries.wf_insert_seria;
end if;

create TRIGGER "wf_insert_seria" before insert on
sguideseries
referencing new as new_name
for each row
begin
	declare v_id_inv integer;
	declare v_values varchar(200);
	declare v_belong_id integer;
	set v_id_inv = get_nextid('inv');
	select id_inv into v_belong_id from sguideseries where seriaId = new_name.parentSeriaId;

	set v_values = convert(varchar(20), v_id_inv)
		+ ', ' + convert(varchar(20), v_belong_id)
		+ ', ''''' + new_name.serianame + ''''''
		+ ', 1'
	;
	
	call insert_host('inv', 'id, belong_id, nm, is_group', v_values);
	set new_name.id_inv=v_id_inv;

/*
	insert into inv (seriaid, parentseriaid, NM, is_group)
	select 
		-new_name.seriaid
		, -new_name.parentseriaid
		, new_name.serianame
		, 1;

  set new_name.id_inv=@@id
*/
end;

if exists (select 1 from systriggers where trigname = 'wf_update_seria' and tname = 'sguideseries') then 
	drop trigger sguideseries.wf_update_seria;
end if;

create TRIGGER "wf_update_seria" before update on
sguideseries
referencing old as old_name new as new_name
for each row
begin
	declare v_id_inv integer;
	declare v_belong_id integer;

	declare v_values varchar(100);
	declare v_fields varchar(200);
	
	set v_id_inv = old_name.id_inv;
	


  if update(serianame) then
	call update_host('inv', 'nm', '''''' + new_name.seriaName + '''''', 'id = ' + convert(varchar(20), v_id_inv));
/*
    update inv as pi set
      nm = new_name.seriaName where
      pi.id = old_name.id_inv
*/
  end if;
  if update(parentSeriaId) then
	select id_inv into v_belong_id from sguideseries where seriaid = new_name.parentseriaId;
	call update_host('inv', 'belong_id', convert(varchar(20), v_belong_id), 'id = ' + convert(varchar(20), v_id_inv));
/*
    update inv as pi set
      belong_id = p.id_Inv
	from sguideseria p
	where
      pi.id = old_name.id_inv
	and p.seriaId = new_name.parentSeriaId
*/
  end if;
end;




if exists (select 1 from systriggers where trigname = 'wf_delete_seria' and tname = 'sGuideSeries') then 
	drop trigger sGuideSeries.wf_delete_seria;
end if;

create TRIGGER "wf_delete_seria" before delete on
sGuideSeries
referencing old as old_name
for each row
begin
	if old_name.id_inv is not null then
		call delete_host('inv', 'id = ' + convert(varchar(20), old_name.id_inv));
	end if;
end;



-------------------------------------------------------------------------
--------------             sGuideNomenk      ----------------------------
-------------------------------------------------------------------------


if exists (select 1 from systriggers where trigname = 'wf_insert_gnomenk' and tname = 'sGuideNomenk') then 
	drop trigger sGuideNomenk.wf_insert_gnomenk;
end if;

create TRIGGER "wf_insert_gnomenk" before insert on
sGuideNomenk
referencing new as new_name
for each row
begin
	declare v_id_inv integer;
	declare v_fields varchar(500);
	declare v_values varchar(2000);
	declare v_belong_id integer;
    declare v_id_edizm1 integer;
    declare v_id_edizm2 integer;
    declare v_id_size integer;

	set v_id_inv = get_nextid('inv');

	select id_inv into v_belong_id from sguideklass where klassId = new_name.KlassId;

	set v_values = convert(varchar(20), v_id_inv)
		+ ', ' + convert(varchar(20), v_belong_id)
		+ ', ''''' + new_name.nomName + ''''''
		+ ', ''''' + new_name.nomnom + ''''''
	;

	set v_fields = 'id, belong_id, nm, nomen';
	if new_name.ed_izmer is not null and length(new_name.ed_izmer) > 0 then
   	  	set v_id_edizm1 = wf_getEdizmId(new_name.ed_izmer);
   	  	set v_fields = v_fields + ', id_edizm2';
   	  	set v_values = v_values + ', '+convert(varchar(20), v_id_edizm1);
   	end if; 

	if new_name.ed_izmer2 is not null and length(new_name.ed_izmer2) > 0 then
	  	set v_id_edizm2 = wf_getEdizmId(new_name.ed_izmer2);
   	  	set v_fields = v_fields + ', id_edizm1';
   	  	set v_values = v_values + ', '+convert(varchar(20), v_id_edizm2);
   	end if; 

	if new_name.size is not null  and length(new_name.size) > 0 then
	  	set v_id_size = wf_getSizeId(new_name.size);
   	  	set v_fields = v_fields + ', id_size';
   	  	set v_values = v_values + ', '+convert(varchar(20), v_id_size);
   	end if; 

	call insert_host('inv', v_fields, v_values);
  set new_name.id_inv=v_id_inv;

end;


if exists (select 1 from systriggers where trigname = 'wf_price_history' and tname = 'sGuideNomenk') then 
	drop trigger sGuideNomenk.wf_price_history;
end if;

create TRIGGER "wf_price_history" before update order 2 on
sGuideNomenk
referencing old as old_name new as new_name
for each row
when (update (cost))
begin
	declare v_changed_by_id tinyint;
	if update(cost) and isnull(old_name.cost, 0) != isnull(new_name.cost, 0)  then
	    begin
			select  managId into v_changed_by_id
			from Guidemanag where manag = @manager;
	    exception when others then
	    	set v_changed_by_id = null;
	    end;
		insert into sPriceHistory (nomnom, cost, change_date, changed_by_id)
		values ( old_name.nomnom, old_name.cost, now(), v_changed_by_id);
	end if;
end;



if exists (select 1 from systriggers where trigname = 'wf_update_gnomenk' and tname = 'sGuideNomenk') then 
	drop trigger sGuideNomenk.wf_update_gnomenk;
end if;

create TRIGGER "wf_update_gnomenk" before update order 1 on
sGuideNomenk
referencing old as old_name new as new_name
for each row
begin
	declare v_id_inv integer;
    declare v_belong_id integer;
    declare v_id_edizm integer;
    declare v_id_size integer;
    declare v_nomName varchar(50);
    declare v_size varchar(30);
    declare v_cod varchar(20);
    declare v_nm varchar(100);

	declare v_values varchar(100);
	declare v_fields varchar(200);
	
	set v_id_inv = old_name.id_inv;
	
  if update(nomnom) then
	call update_host('inv', 'nomen', '''''' + new_name.nomnom + '''''', 'id = ' + convert(varchar(20), v_id_inv));
  end if;

  if update(ed_Izmer) then
  	set v_id_edizm = wf_getEdizmId(new_name.ed_izmer);
--	select id_edizm into v_ed_izm from edizm where e.name = new_name.ed_izmer;
	call update_host('inv', 'id_edizm2', convert(varchar(20), v_id_edizm), 'id = ' + convert(varchar(20), v_id_inv));
  end if;

  if update(ed_Izmer2) then
  	set v_id_edizm = wf_getEdizmId(new_name.ed_izmer2);
--	select id_edizm into v_ed_izm from edizm where e.name = new_name.ed_izmer;
	call update_host('inv', 'id_edizm1', convert(varchar(20), v_id_edizm), 'id = ' + convert(varchar(20), v_id_inv));
  end if;
  
  if update(klassId) then
	select id_inv into v_belong_id from sguideklass where klassid = new_name.klassId;
	call update_host('inv', 'belong_id', convert(varchar(20), v_belong_id), 'id = ' + convert(varchar(20), v_id_inv));
  end if;
  
  if update(size) or update (cod) or update(nomName) then
  	if (new_name.nomName != old_name.nomName) then
  		set v_nomName = new_name.nomName;
  	else 
  		set v_nomName = old_name.nomName;
  	end if;

  	if (new_name.cod != old_name.cod) then
  		set v_cod = new_name.cod;
  	else 
  		set v_cod = old_name.cod;
  	end if;

  	if (new_name.size != old_name.size) then
  		set v_size = new_name.size;
	  	set v_id_size = wf_getSizeId(new_name.size);
		call update_host('inv', 'id_size', convert(varchar(20), v_id_size), 'id = ' + convert(varchar(20), v_id_inv));
  	else 
  		set v_size = old_name.size;
  	end if;


	set v_nm = wf_make_invnm (v_nomname, v_size, v_cod);
	call update_host('inv', 'nm', '''''' + v_nm + '''''', 'id = ' + convert(varchar(20), v_id_inv));

  end if;


end;


if exists (select 1 from systriggers where trigname = 'wf_delete_gnomenk' and tname = 'sGuideNomenk') then 
	drop trigger sGuideNomenk.wf_delete_gnomenk;
end if;

create TRIGGER "wf_delete_gnomenk" before delete on
sGuideNomenk
referencing old as old_name
for each row
begin
	if old_name.id_inv is not null then
		call delete_host('inv', 'id = ' + convert(varchar(20), old_name.id_inv));
	end if;
end;

--------------------------------------------------------------------------
--------------             sGuideProducts      ----------------------------
--------------------------------------------------------------------------


if exists (select 1 from systriggers where trigname = 'wf_insert_gproduct' and tname = 'sguideproducts') then 
	drop trigger sguideproducts.wf_insert_gproduct;
end if;

create TRIGGER "wf_insert_gproduct" before insert on
sguideproducts
referencing new as new_name
for each row
begin
	declare v_id_inv integer;
	declare v_fields varchar(500);
	declare v_values varchar(2000);
	declare v_belong_id integer;
    declare v_id_edizm1 integer;
    declare v_id_size integer;
    declare v_nm varchar(102);


	set v_id_inv = get_nextid('inv');

	select id_inv into v_belong_id from sguideseries where seriaId = new_name.prSeriaId;
  	set v_id_edizm1 = wf_getEdizmId('шт.');

  set v_fields = 
  	  ' id'
  	+ ',belong_id'
  	+ ',nomen'
    + ',nm'
    + ',prc1'
    + ',is_compl'
    + ', id_edizm1'
	;

	set v_nm = wf_make_invnm (new_name.prDescript, new_name.prSize, new_name.prName);

	set v_values = 
				 convert(varchar(20), v_id_inv)
		+ ', ' + convert(varchar(20), v_belong_id)
		+ ', ''''' + new_name.prName + ''''''
		+ ', ''''' + v_nm + ''''''
		+ ', ' + convert(varchar(20), new_name.cena4)
		+ ', 1'
   	  	+ ', '+convert(varchar(20), v_id_edizm1);
	;


	if new_name.prsize is not null and length(new_name.prsize) > 0 then
	  	set v_id_size = wf_getEdizmId(new_name.prsize);
   	  	set v_fields = v_fields + ', id_size';
   	  	set v_values = v_values + ', '+convert(varchar(20), v_id_size);
   	end if; 

	call insert_host('inv', v_fields, v_values);
  set new_name.id_inv=v_id_inv;
	

end;



if exists (select 1 from systriggers where trigname = 'wf_update_gproducts' and tname = 'sGuideProducts') then 
	drop trigger sGuideProducts.wf_update_gproducts;
end if;

create TRIGGER "wf_update_gproducts" before update on
sGuideProducts
referencing old as old_name new as new_name
for each row
begin
	declare v_id_inv integer;
    declare v_belong_id integer;
    declare v_id_edizm integer;
    declare v_id_size integer;
    declare v_prDescript varchar(50);
    declare v_prSize varchar(30);
    declare v_prName varchar(20);
    declare v_nm varchar(102);
    declare is_variant integer;

	declare v_values varchar(100);
	declare v_fields varchar(200);
	
	set v_id_inv = old_name.id_inv;


  if update(prSize) or update(prName) or update (prDescript) then

	select 1 into is_variant from svariantpower vp where vp.productid = old_name.prId;

	if (new_name.prDescript != old_name.prDescript) then
		set v_prDescript = new_name.prDescript;
	else 
		set v_prDescript = old_name.prDescript;
	end if;
  
	if (new_name.prName != old_name.prName) then
		set v_prName = new_name.prName;
		call update_host('inv', 'nomen', '''''' + new_name.prName + '''''', 'id = ' + convert(varchar(20), v_id_inv));
		if is_variant is not null then
			call update_host('inv', 'nomen', '''''' + new_name.prName + '''''', 'belong_id = ' + convert(varchar(20), v_id_inv));
		end if;
	else 
		set v_prName = old_name.prName;
	end if;
  
	if (new_name.prSize != old_name.prSize) then
		set v_prSize = new_name.prSize;
		set v_id_size = wf_getSizeId(v_prSize);
		call update_host('inv', 'id_size', convert(varchar(20), v_id_size), 'id = ' + convert(varchar(20), v_id_inv));
		if is_variant is not null then
			call update_host('inv', 'id_size', convert(varchar(20), v_id_size), 'belong_id = ' + convert(varchar(20), v_id_inv));
		end if;
	else 
		set v_prSize = old_name.prSize;
	end if;
  
  
	set v_nm = wf_make_invnm (v_prDescript, v_prSize, v_prName);
	call update_host('inv', 'nm', '''''' + v_nm + '''''', 'id = ' + convert(varchar(20), v_id_inv));
	if is_variant is not null then
		
		for aCursor as a dynamic scroll cursor for
			select 
				  xprext as r_xprext
				, id_inv as r_id_inv_variant
			from sguidecomplect g
			where productid = old_name.prid
		do
			set v_nm = wf_make_variant_nm (
				  v_prDescript
				, v_prSize
				, v_prName
				, r_xprext
			);
			call update_host('inv', 'nm', '''''' + v_nm + '''''', 'id = ' + convert(varchar(20), r_id_inv_variant));
			call update_host('inv', 'nomen', '''''' + v_prName + '''''', 'id = ' + convert(varchar(20), r_id_inv_variant));

		end for;
	end if;

  end if;

/*
  if update(prName) then
	call update_host('inv', 'nomen', '''''' + new_name.prName + '''''', 'id = ' + convert(varchar(20), v_id_inv));
  end if;

  if update(prDescript) then
	call update_host('inv', 'nm', '''''' + new_name.prDescript + '''''', 'id = ' + convert(varchar(20), v_id_inv));
  end if;

  if update(prsize) then
  	set v_id_size = wf_getSizeId(new_name.prsize);
		call update_host('inv', 'id_size', convert(varchar(20), v_id_size), 'id = ' + convert(varchar(20), v_id_inv));
  end if;
*/


  if update(seriaId) then
	select id_inv into v_belong_id from sguideseries where seriaId = new_name.prSeriaId;
	call update_host('inv', 'belong_id', convert(varchar(20), v_belong_id), 'id = ' + convert(varchar(20), v_id_inv));
  end if;
  
end;



if exists (select 1 from systriggers where trigname = 'wf_delete_gproducts' and tname = 'sGuideProducts') then 
	drop trigger sGuideProducts.wf_delete_gproducts;
end if;

create TRIGGER "wf_delete_gproducts" before delete on
sGuideProducts
referencing old as old_name
for each row
begin
	if old_name.id_inv is not null then
		call delete_host('inv', 'id = ' + convert(varchar(20), old_name.id_inv));
	end if;
end;


----------------------------------------------------------------------
--------------             sProducts      ----------------------------
----------------------------------------------------------------------



if exists (select 1 from systriggers where trigname = 'wf_insert_product' and tname = 'sProducts') then 
	drop trigger sProducts.wf_insert_product;
end if;

create TRIGGER "wf_insert_product" before insert order 1 on
sProducts
referencing new as new_name
for each row
begin

  declare v_table_name varchar(30);
  declare v_values varchar(100);
  declare v_fields varchar(200);
  
  declare v_id_inv integer; -- id номенклатуры
  declare v_id_belong_inv integer; -- id изделия
  declare v_id_compl integer; -- backref
  
  declare is_variant integer; -- проверка того, что изделие простое
  
  declare v_id_edizm integer;
  declare v_edizm varchar(50);
  
  
  update sGuideVariant as gv set c = c+1 where gv.productid = new_name.productId and gv.xgroup = new_name.xgroup;
  if @@rowcount = 0 then
    insert into sGuideVariant(c,productid,xgroup) values(
      1,new_name.productId,new_name.xgroup)
  end if;
  
  select numgroup into is_variant from svariantpower where productid = new_name.productid;

  --if (is_variant is null) then
	//Грузим комплектацию 
    // простое (не вариантное) (пока!) изделие
	set v_table_name = 'compl';
	set v_id_compl = get_nextId (v_table_name);
	select id_inv, ed_izmer into v_id_Inv, v_edizm from sguidenomenk where nomnom = new_name.nomNom;
	set v_id_edizm = wf_getEdizmId (v_edizm);

	select id_inv into v_id_belong_inv from sguideproducts where prid = new_name.productId;
	
	set v_fields ='id'
		+ ', id_inv'
		+ ', id_inv_belong'
		+ ', id_edizm'
		+ ', kol'
		;
	
	set v_values =
			 convert(varchar(20), v_id_compl )
			+ ', ' + convert(varchar(20), v_id_inv)
			+ ', ' + convert(varchar(20), v_id_belong_inv)
			+ ', ' + convert(varchar(20), v_id_edizm)
			+ ', ' + convert(varchar(20), new_name.quantity)
		;	

	call insert_host (v_table_name, v_fields, v_values);
	set new_name.id_compl = v_id_compl;
  --end if;
/*
  insert into compl (id_inv, id_inv_belong, id_edizm, kol)
	select gn.id_inv, gp.id_inv, wf_getEdizmId (gn.ed_izmer), new_name.quantity
  from sguideproducts gp
  join sguidenomenk gn on gn.nomNom = new_name.nomNom
  where gp.prid = new_name.productId;
*/

end;



if exists (select 1 from systriggers where trigname = 'wf_update_product' and tname = 'sProducts') then 
	drop trigger sProducts.wf_update_product;
end if;

create TRIGGER "wf_update_product" before update on
sProducts
referencing old as old_name new as new_name
for each row
begin
  declare namedFromAfter integer;
  
  if update(xgroup) then
  	update sGuideVariant as gv set c = c-1 where gv.productid = old_name.productId and gv.xgroup = old_name.xgroup;
	select c into namedFromAfter from sGuideVariant gv where gv.productid = old_name.productId and gv.xgroup = old_name.xgroup;
	if namedFromAfter = 0 then
		delete from sGuideVariant where productid = old_name.productId and xgroup = old_name.xgroup;
	end if;

	update sGuideVariant as gv set c = c+1 where gv.productid = old_name.productId and gv.xgroup = new_name.xgroup;
	if @@rowcount = 0 then
	 		insert into sGuideVariant (c, productid, xgroup) 
			values( 1, old_name.productId, new_name.xgroup);
	end if;
  	
  
  end if;

  if update (quantity) then
	call update_host('compl', 'kol', convert(varchar(20), new_name.quantity), 'id = ' + convert(varchar(20), old_name.id_compl))
/*
 	update compl c set kol = new_name.quantity
  	from sguideproducts gp  
  	join sguidenomenk gn on gn.nomNom = old_name.nomNom
  	where gp.prid = old_name.productid 
  	and c.id_inv = gn.id_inv and c.id_inv_belong = gp.id_inv;
*/
  end if;

end;


if exists (select 1 from systriggers where trigname = 'wf_delete_product' and tname = 'sProducts') then 
	drop trigger sProducts.wf_delete_product;
end if;

create TRIGGER "wf_delete_product" after delete on
sProducts
referencing old as old_name
for each row
begin
    declare namedFromAfter integer;
  	update sGuideVariant as gv set c = c-1 where gv.productid = old_name.productId and gv.xgroup = old_name.xgroup;
	select c into namedFromAfter from sGuideVariant gv where gv.productid = old_name.productId and gv.xgroup = old_name.xgroup;
	if namedFromAfter <= 0 then
		delete from sGuideVariant where productid = old_name.productId and xgroup = old_name.xgroup;
	end if;
	if old_name.id_compl is not null then
		call delete_host('compl', 'id = ' + convert(varchar(20), old_name.id_compl));
	end if;
end;
----------------------------------------------------------------------
--------------             sGuideVariant      ------------------------
----------------------------------------------------------------------



if exists (select 1 from systriggers where trigname = 'wf_insert_gvariant' and tname = 'sGuideVariant') then 
	drop trigger sGuideVariant.wf_insert_gvariant;
end if;

create TRIGGER "wf_insert_gvariant" after insert on
sGuideVariant
referencing new as new_name
for each row
begin
	-- вроде бы ничего не нужно делать
	-- в штатном режиме при добавлении номенклатуры к изделию
	-- добавиться может только либо строка с пустой xgroup
	-- либо строка со значением счетчика, равной 1
	-- И в том и другом случае состояние вариантности не меняется.
end;


if exists (select 1 from systriggers where trigname = 'wf_update_gvariant' and tname = 'sGuideVariant') then 
	drop trigger sGuideVariant.wf_update_gvariant;
end if;


create TRIGGER "wf_update_gvariant" before update on
sGuideVariant
referencing old as old_name new as new_name
for each row
begin
	declare v_power integer;
	declare v_fixgroups integer;
	
	if update(c) then
		if old_name.xgroup != '' then
			select c into v_fixgroups from sGuideVariant where productId = old_name.productid and xgroup = '';
			select numgroup into v_power from svariantpower where productid = old_name.productid;
			if old_name.c = 1 and new_name.c = 2 then
				update svariantpower set numgroup = numgroup + 1 where productid = old_name.productid;
				if @@rowcount = 0 then
					-- изделие становится вариантным
					insert into svariantpower (numgroup, productid, fixgroups)
					values (1, old_name.productid, v_fixgroups);
				end if;
			elseif old_name.c = 2 and new_name.c = 1 then
				update svariantpower set numgroup = numgroup - 1 where productid = old_name.productid;
				select numgroup into v_power from svariantpower where productid = old_name.productid;
				if v_power = 0 then
					-- изделие перестает быть вариантным
					delete from svariantpower where productid = old_name.productid;
				end if;
			end if;
		else
			-- апдейтим количество фиксированных компонент (если конечно изделие вариантное)
			update svariantpower set fixgroups = new_name.c where productid = old_name.productid;
		end if;
		
	end if;
	
end;


if exists (select 1 from systriggers where trigname = 'wf_delete_gvariant' and tname = 'sGuideVariant') then 
	drop trigger sGuideVariant.wf_delete_gvariant;
end if;

create TRIGGER "wf_delete_gvariant" after delete on
sGuideVariant
referencing old as old_name
for each row
begin
end;

----------------------------------------------------------------------
--------------             sVariantPower      ------------------------
----------------------------------------------------------------------



if exists (select 1 from systriggers where trigname = 'wf_insert_vpower' and tname = 'sVariantPower') then 
	drop trigger sVariantPower.wf_insert_vpower;
end if;

create TRIGGER "wf_insert_vpower" after insert on
sVariantPower
referencing new as new_name
for each row
begin
	declare v_id_inv integer;
	select id_inv into v_id_inv from sguideproducts where prid = new_name.productid;
	call update_host('inv', 'is_group', '1', 'id = ' + convert(varchar(20), v_id_inv));
	call update_host('inv', 'is_compl', '0', 'id = ' + convert(varchar(20), v_id_inv));
end;


if exists (select 1 from systriggers where trigname = 'wf_delete_vpower' and tname = 'sVariantPower') then 
	drop trigger sVariantPower.wf_delete_vpower;
end if;

create TRIGGER "wf_delete_vpower" after delete on
sVariantPower
referencing old as old_name
for each row
begin
	declare v_id_inv integer;
	select id_inv into v_id_inv from sguideproducts where prid = old_name.productid;
	call update_host('inv', 'is_group', '0', 'id = ' + convert(varchar(20), v_id_inv));
	call update_host('inv', 'is_compl', '1', 'id = ' + convert(varchar(20), v_id_inv));
end;

----------------------------------------------------------------------
--------------             sdmc      ------------------------
----------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_sdmc_income_bu' and tname = 'sdmc') then 
	drop trigger sdmc.wf_sdmc_income_bu;
end if;

create 
	trigger wf_sdmc_income_bu before update  on 
sdmc
referencing new as new_name old as old_name
for each row
begin
	declare v_id_mat integer;
	declare v_perList float;
	declare v_curr float;
	declare v_id_jmat integer;
	declare v_summa float;
	declare v_summav float;
	declare v_quant float;

	if update(quant) --and isnull(new_name.quant, 0.0) != isnull(old_name.quant, 0.0) 
	then
		set v_id_mat = old_name.id_mat;
		select perList into v_perList from sGuideNomenk where nomnom = old_name.nomNom;
		if v_id_mat is not null then
			select id_jmat into v_id_jmat from sdocs where numdoc = old_name.numdoc and numext = old_name.numext;
			set v_quant = new_name.quant/v_perList;

			call change_mat_qty_stime(v_id_mat, v_quant);
		end if;
	end if;
/*
	-- Не нужно, потому что в интерфейсе stime нельзя заменить номенклатуру.
	-- Можно только сначала удалить позицию, а потом завести новую.
	if update (nomnom) and isnull(new_name.nomnom, '') != isnull(old_name.nomnom, '') then
		set v_id_mat = old_name.id_mat;
		if v_id_mat is not null then
			select id_inv into v_id_inv from sguidenomenk where nomnom = new_name.nomnom;
			if v_id_inv is not null then
				call update_remote('stime', 'mat', 'id_inv', '''''' + v_id_inv + '''''', 'id = ' + convert(varchar(20), v_id_mat)); 
			end if;
		end if;
	end if;
*/
end;


if exists (select 1 from systriggers where trigname = 'wf_delete_sdmc' and tname = 'sdmc') then 
	drop trigger sdmc.wf_delete_sdmc;
end if;

create TRIGGER wf_delete_sdmc before delete on
sdmc
referencing old as old_name
for each row
begin
	declare remoteServer varchar(32);
	declare no_echo integer;
	set no_echo = 0;


  	begin
  		message '@stime_sdmc = ', @stime_sdmc to log;
		select @stime_sdmc into no_echo; 
	exception 
		when other then
			message 'Exception! no_echo = ' + convert(varchar(20), no_echo) to log;
			set no_echo = 0;
	end;

	--message 'trigger sdmc.wf_delete_sdmc::no_echo = ' + convert(varchar(20), no_echo) to log;
	if no_echo = 1 then
		return;
	end if;




	if (old_name.id_mat is not null) then
		call block_remote('stime', @@servername, 'mat');
		call delete_remote('stime', 'mat', 'id = ' + convert(varchar(20), old_name.id_mat));
		call unblock_remote('stime', @@servername, 'mat');
	end if;

	--message 'old_name.id_mat = ', old_name.id_mat to client;

	select sysname into remoteServer 
	from  guideventure v 
	join orders o on o.ventureId = v.ventureId and v.standalone = 0 and o.numorder = old_name.numDoc;

	--message 'remoteServer = ', remoteServer to client;

	if remoteServer is not null and remoteServer != 'stime' then
		call block_remote(remoteServer, @@servername, 'mat');
		call delete_remote(remoteServer, 'mat', 'id = ' + convert(varchar(20), old_name.id_mat));
		call unblock_remote(remoteServer, @@servername, 'mat');
	end if;

end;


if exists (select 1 from systriggers where trigname = 'wf_sdmc_outcome_bi' and tname = 'sdmc') then 
	drop trigger sdmc.wf_sdmc_outcome_bi;
end if;

create 
	trigger wf_sdmc_outcome_bi before insert order 3 on 
sdmc
referencing new as new_name
for each row
begin
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_mat_nu varchar(20);
	declare v_currency_rate float;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_id_inv integer;
	declare v_id_source integer;
	declare v_id_dest integer;
	declare v_cost float;
	declare v_quant float;
	declare v_perList float;
	declare sync char(1);
	declare no_echo integer;

	set no_echo = 0;

	--message 'no_echo = ' + convert(varchar(20), no_echo) to log;

  	begin
  		message '@stime_sdmc = ', @stime_sdmc to log;
		select @stime_sdmc into no_echo; 
	exception 
		when other then
			--message 'Exception! no_echo = ' + convert(varchar(20), no_echo) to log;
			set no_echo = 0;
	end;

	if no_echo = 1 then
		return;
	end if;

	select id_jmat, s.id_voc_names, d.id_voc_names
	into v_id_jmat, v_id_source, v_id_dest
	from sdocs n 
		join sguidesource s on s.sourceid = n.sourid 
		join sguidesource d on d.sourceid = n.destid
	where n.numdoc = new_name.numdoc and n.numext = new_name.numext;

	if v_id_jmat is null then
		return;
	end if;

	set v_id_mat = get_nextid('mat');

	set v_id_currency = system_currency();
	call slave_currency_rate_stime(v_datev, v_currency_rate);
	call slave_select_stime(v_mat_nu, 'mat', 'max(nu)', 'id_jmat = ' + convert(varchar(20), v_id_jmat));
	
	set v_mat_nu = convert(varchar(20), convert(integer, isnull(v_mat_nu, 0)) + 1);

	select 
		id_inv
		, cost 
		, perList
	into 
		v_id_inv
		, v_cost 
		, v_perList
	from sguidenomenk 
	where nomnom = new_name.nomnom;


	set v_quant = new_name.quant; -- / v_perList;

--		call call_host('block_table', 'sync, ''prior'', ''mat''');
		call block_remote('stime', @@servername, 'mat');
	
		call wf_insert_mat (
			'stime'
			,v_id_mat
			,v_Id_jmat
			,v_id_inv
			,v_mat_nu
			,v_quant 
			,v_cost
			,v_currency_rate
			,v_id_source
			,v_id_dest
			,v_perList
		);

		set new_name.id_mat = v_id_mat;
		call unblock_remote('stime', @@servername, 'mat');
end;


----------------------------------------------------------------------
--------------                 sdocs          ------------------------
----------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_delete_sdocs' and tname = 'sdocs') then 
	drop trigger sdocs.wf_delete_sdocs;
end if;

create TRIGGER wf_delete_sdocs before delete on
sdocs
referencing old as old_name
for each row
begin
	declare remoteServer varchar(32);
	declare no_echo integer;

	set no_echo = 0;

  	begin
  		message '@stime_sdocs = ', @stime_sdocs to log;
		select @stime_sdocs into no_echo; 
	exception 
		when other then
			set no_echo = 0;
	end;

	if no_echo = 1 then
		return;
	end if;



	if (old_name.id_jmat is not null) then
		call block_remote('stime', @@servername, 'jmat');
		call block_remote('stime', @@servername, 'mat');
		call delete_remote('stime', 'jmat', 'id = ' + convert(varchar(20), old_name.id_jmat));
		call unblock_remote('stime', @@servername, 'jmat');
		call unblock_remote('stime', @@servername, 'mat');
	end if;

	select sysname into remoteServer 
	from  guideventure v 
	join orders o on o.ventureId = v.ventureId and v.standalone = 0 and o.numorder = old_name.numDoc;

--	message 'remoteServer = ', remoteServer to client;
	if remoteServer is not null and remoteServer != 'stime' then
		call block_remote(remoteServer, @@servername, 'jmat');
		call block_remote(remoteServer, @@servername, 'mat');
		call delete_remote(remoteServer, 'jmat', 'id = ' + convert(varchar(20), old_name.id_jmat));
		call unblock_remote(remoteServer, @@servername, 'jmat');
		call unblock_remote(remoteServer, @@servername, 'mat');
	end if;
end;


if exists (select 1 from systriggers where trigname = 'wf_set_numdoc' and tname = 'sdocs') then 
	drop trigger sdocs.wf_set_numdoc;
end if;

create 
	trigger wf_set_numdoc before insert order 1 on 
sdocs
referencing new as new_name
for each row
when (new_name.numdoc = 0 or new_name.numdoc is null)
begin
	set new_name.numdoc = wf_next_numdoc();
end;


if exists (select 1 from systriggers where trigname = 'wf_insert_income' and tname = 'sdocs') then 
	drop trigger sdocs.wf_insert_income;
end if;


create 
	trigger wf_insert_income before insert order 2 on 
sdocs
referencing new as new_name
for each row
when (new_name.numext = 255)
begin
	
end;



if exists (select 1 from systriggers where trigname = 'wf_sdocs_outcome_bi' and tname = 'sdocs') then 
	drop trigger sdocs.wf_sdocs_outcome_bi;
end if;

create 
	trigger wf_sdocs_outcome_bi before insert order 3 on 
sdocs
referencing new as new_name
for each row
--when (new_name.numext <= 254)
begin
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_jmat_nu varchar(20);
	declare v_currency_rate float;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_id_source integer;
	declare v_id_dest integer;
	declare v_osn varchar(100);
	declare v_id_guide_jmat integer;
	declare v_currency_iso varchar(10);
	declare no_echo integer;

	set no_echo = 0;

  	begin
  		message '@stime_sdocs = ', @stime_sdocs to log;
		select @stime_sdocs into no_echo; 
	exception 
		when other then
			set no_echo = 0;
	end;

	if no_echo = 1 then
		return;
	end if;

	message 'before if new_name.numext = 254 then' to client;
	if new_name.numext = 254 then
		set v_id_guide_jmat = 1220;
	elseif new_name.numext = 255 and new_name.sourId is not null then
		select 
			  isnull(c.id_guide, ru.id_guide) 
			, isnull(c.id_currency, ru.id_currency) 
			, isnull(c.currency_iso, ru.currency_iso)
		into v_id_guide_jmat, v_id_currency, v_currency_iso
		from sguideSource s
		join GuideCurrency ru on ru.currency_iso = 'RUR'
		left join GuideCurrency c on c.currency_iso = s.currency_iso
		where s.sourceId = new_name.sourId;
	else 
		set v_id_guide_jmat = 1210;
	end if;

	set v_id_jmat = get_nextid('jmat');
	

	message 'before if isnull(v_currency_iso, ...' to client;
	if isnull(v_currency_iso, 'RUR') = 'RUR' then
		set v_id_currency = system_currency();
	end if;

	call slave_currency_rate_stime(v_datev, v_currency_rate, null, v_id_currency);

	set v_jmat_nu = new_name.numdoc;
	select id_voc_names into v_id_source from sguidesource where sourceid = new_name.sourid;
	select id_voc_names into v_id_dest from sguidesource where sourceid = new_name.destid;
	set v_osn = '[Prior: '+ convert(varchar(20), new_name.numdoc) +']';
    
	call wf_insert_jmat (
		'stime'
		,v_id_guide_jmat
		,v_id_jmat
		,now() --v_jmat_date
		,v_jmat_nu
		,v_osn
		,v_id_currency
		,v_datev
		,v_currency_rate
		,v_id_source
		,v_id_dest
	);
	set new_name.id_jmat = v_id_jmat;



end;

if exists (select 1 from systriggers where trigname = 'wf_sdocs_outcome_bu' and tname = 'sdocs') then 
	drop trigger sdocs.wf_sdocs_outcome_bu;
end if;

create 
	trigger wf_sdocs_outcome_bu before update on 
sdocs
referencing new as new_name old as old_name
for each row
--when (old_name.numext = 254)
begin
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_jmat_nu varchar(20);
	declare v_currency_rate float;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_id_source integer;
	declare v_id_dest integer;
	declare v_osn varchar(100);

	declare v_id_guide integer;
	declare v_tp1 integer;
	declare v_tp2 integer;
	declare v_tp3 integer;
	declare v_tp4 integer;
	declare v_currency_iso varchar(20);
	
	declare no_echo integer;

	set no_echo = 0;

	if old_name.id_jmat is null then 
		return;
	end if;
	
	begin
  		message '@stime_sdocs = ', @stime_sdocs to log;
		select @stime_sdocs into no_echo;
	exception 
		when other then
			message '.... exception ' to log;
			set no_echo = 0;
	end;

	if no_echo = 1 then
		return;
	end if;


	message 'sdocs.wf_sdocs_outcome_bu old_id = ' to log;
	call block_remote ('stime', @@servername, 'jmat');

	if update(sourid) then
		select id_voc_names into v_id_source from sguidesource where sourceid = new_name.sourid;
		if v_Id_source is not null then
			call update_remote('stime', 'jmat', 'id_s', convert(varchar(20), v_id_source), 'id = ' + convert(varchar(20), old_name.id_jmat));
		end if;
		if old_name.numext = 255 and old_name.id_jmat is not null then
			select 
				  isnull(c.id_guide, ru.id_guide)
				, isnull(c.id_currency, ru.id_currency) 
				, isnull(c.currency_iso, ru.currency_iso)
			into v_id_guide, v_id_currency, v_currency_iso
			from sguideSource s
			join GuideCurrency ru on ru.currency_iso = 'RUR'
			left join GuideCurrency c on c.currency_iso = s.currency_iso
			where s.sourceId = new_name.sourId;
	    
			if isnull(v_currency_iso, 'RUR') = 'RUR' then
				set v_id_currency = system_currency();
			end if;
			call gualify_guide(v_id_guide, v_tp1, v_tp2, v_tp3, v_tp4);
			call order_import_stime(
				  old_name.id_jmat
				, v_id_currency
				, v_id_guide
				, v_tp1
				, v_tp2
				, v_tp3
				, v_tp4
			);
		end if;

	end if;
	if update(destid) then
		select id_voc_names into v_id_dest from sguidesource where sourceid = new_name.destid;
		call update_remote('stime', 'jmat', 'id_d', convert(varchar(20), v_id_dest), 'id = ' + convert(varchar(20), old_name.id_jmat));
	end if;
	if update(xDate) then
		call update_remote('stime', 'jmat', 'dat', '''''' + convert(varchar(20), new_name.xDate) + '''''', 'id = ' + convert(varchar(20), old_name.id_jmat));
	end if;

	--if update(note) then
		-- set v_osn = '[Prior: '+ new_name.note +']';
		-- пришлось отключить из-за ошибки при установки 
		-- признака предприятия в приходной накладной
		-- call update_remote ('stime', 'jmat', 'osn', '''' +v_osn + '''', 'id = ' + convert(varchar(20), old_name.id_jmat));
	--end if;
	call unblock_remote ('stime', @@servername, 'jmat');
end;



-------------------------------------------------------------------------
--------------             Orders      ----------------------------
-------------------------------------------------------------------------


if exists (select 1 from systriggers where trigname = 'wf_insert_orders' and tname = 'Orders') then 
	drop trigger Orders.wf_insert_orders;
end if;

create TRIGGER "wf_insert_orders" before insert on
Orders
referencing new as new_name
for each row
begin
end;


if exists (select 1 from systriggers where trigname = 'wf_update_orders' and tname = 'Orders') then 
	drop trigger Orders.wf_update_orders;
end if;

create TRIGGER wf_update_orders before update order 1 on
Orders
referencing old as old_name new as new_name
for each row
begin
	declare remoteServerOld varchar(32);
	declare remoteServerNew varchar(32);
	declare v_fields varchar(255);
	declare v_values varchar(2000);
	declare v_nu_jscet integer;
	declare r_nu varchar(50);
	declare r_id integer;
--	declare v_firm_id integer;
	declare v_invCode varchar(10);
	declare v_id_dest integer;
	declare v_id_schef integer;
	declare v_id_bux integer;
	declare v_id_bank integer;
	declare v_datev varchar(20);
	declare v_id_cur integer;
	declare v_currency_rate float;
	declare v_order_date varchar(20);
	declare v_check_count integer; 
	declare v_id_jscet integer;
	declare v_id_scet integer;
	declare v_id_inv integer;
	declare v_numorder integer;
	declare v_updated integer;

	declare sync char(1);

	select sysname, invCode into remoteServerOld, v_invcode from GuideVenture where ventureId = old_name.ventureId;

	if update(invoice) and remoteServerOld is not null then
		call update_remote(remoteServerOld, 'jscet', 'nu'
				, convert(varchar(20), extract_invoice_number(new_name.invoice, v_invCode))
				, 'id = ' + convert(varchar(20), old_name.id_jscet)
		);
	end if;


	if update(ventureId) then
		if new_name.ventureId = 0 then
			set new_name.ventureid = null;
		end if;
		if isnull(old_name.ventureId, 0) != isnull(new_name.ventureId, 0) then
			if remoteServerOld is not null then
				call delete_remote(remoteServerOld, 'jscet', 'id = ' + convert(varchar(20), old_name.id_jscet));
				call delete_remote(remoteServerOld, 'scet', 'id_jmat = ' + convert(varchar(20), old_name.id_jscet));
				set new_name.invoice = 'счет ?';
				set new_name.id_bill = null;
			end if;
		
			select sysname, invCode into remoteServerNew, v_invcode from GuideVenture where ventureId = new_name.ventureId;
			if remoteServerNew is not null then
		
				set v_numOrder = old_name.numOrder;
				select id_voc_names into v_id_dest from guidefirms where firmid = old_name.firmId;
--				set v_firm_id = old_name.firmId;
				call put_jscet(r_id, v_nu_jscet, remoteServerNew, v_numOrder, v_id_dest, old_name.invoice);
		
				set new_name.id_jscet = r_id;
				set new_name.invoice = v_invCode + convert(varchar(20), v_nu_jscet);
				call wf_set_invoice_detail(remoteServerNew, r_id, new_name.numOrder, v_order_date);
			end if;
		end if;
	end if;
	if update (firmId) and (old_name.id_bill is null or old_name.id_bill = 0) then
		
		select sysname into remoteServerOld from GuideVenture where ventureId = old_name.ventureId;
		if remoteServerOld is not null then
			select id_voc_names into v_id_dest from guideFirms where firmId = new_name.firmId;
			call block_remote(remoteServerOld, @@servername, 'jscet');
			call update_remote(remoteServerOld, 'jscet', 'id_d', convert(varchar(20), v_id_dest ), 'id = ' + convert(varchar(20), old_name.id_jscet));
			call update_remote(remoteServerOld, 'jscet', 'id_d_cargo', convert(varchar(20), v_id_dest ), 'id = ' + convert(varchar(20), old_name.id_jscet));
			call unblock_remote(remoteServerOld, @@servername, 'jscet');
		end if;
	end if;
	if update (ordered) then

		set v_id_jscet = old_name.id_jscet;
	
		if remoteServerOld is not null and v_id_jscet is not null then
//			message 'remoteServerOld = ', remoteServerOld to client;
			-- Заказ, который имеет ссылки в бух.базах интеграции
			-- т.е. уже назначен той, иди другой фирме

			-- отследить заказ без предметов
			-- сначала проверяем, что он действительно без них
			select count(*) into v_check_count from xpredmetybynomenk where numorder = old_name.numorder;
			if v_check_count > 0 then
				-- заказ с предметами -> ничего не делаем
				return;
			end if;
	    
			select count(*) into v_check_count from xpredmetybyizdelia where numorder = old_name.numorder;
			if v_check_count > 0 then
				-- заказ с предметами -> ничего не делаем
				return;
			end if;
	    
			-- ищем товар под названием "услуга"
			select id_inv into v_id_inv from sGuideNomenk where nomNom = 'УСЛ';

			-- сначала исходим из того, что такая услуга уже есть.
			-- это может произойти при изменении стоимости закакза.

			if abs(new_name.ordered) < 0.001 then
				call delete_remote(remoteServerOld, 'scet'
					, 'id_jmat = ' + convert(varchar(20), v_id_jscet) + ' and id_inv = ' + convert(varchar(20), v_id_inv)
				);
				return;
			end if;
			set v_updated = update_count_remote(
				remoteServerOld
				,'scet', 'summa_salev'
				, convert(varchar(20), new_name.ordered)
				, 'id_jmat = ' + convert(varchar(20), v_id_jscet) + ' and id_inv = ' + convert(varchar(20), v_id_inv)
			);
			execute immediate 'call slave_currency_rate_' + remoteServerOld + '(v_datev, v_currency_rate, v_order_date, v_id_cur )';
//			message 'v_currency_rate = ', v_currency_rate to client;
			set v_updated = update_count_remote(
				remoteServerOld
				,'scet', 'summa_sale'
				, convert(varchar(20), new_name.ordered * v_currency_rate)
				, 'id_jmat = ' + convert(varchar(20), v_id_jscet) + ' and id_inv = ' + convert(varchar(20), v_id_inv)
			);


//			message 'v_updated = ', v_updated to client;
			if v_updated > 0 then
				-- именно такой случай
				return;
			end if;

//			message 'v_id_jscet = ',v_id_jscet  to client;
//			message 'v_id_inv = ', v_id_inv to client;

			-- первый раз меням это поле => нужно добавить
			set v_id_scet = 
				wf_insert_scet(
					remoteServerOld
					, v_id_jscet
					, v_id_inv
					, 1
					, new_name.ordered
					, old_name.indate
				);
		end if;


	end if;
end;


if exists (select 1 from systriggers where trigname = 'last_modified' and tname = 'orders') then 
	drop trigger orders.last_modified;
end if;

create TRIGGER last_modified before update order 2 on 
orders
referencing old as old_name new as new_name
for each row
begin
	if not update(rowLock) and not update(numorder) and not update(lastModified) and not update(id_bill) then
		set new_name.lastModified = now();
	end if;
end;




if exists (select 1 from systriggers where trigname = 'wf_delete_orders' and tname = 'Orders') then 
	drop trigger Orders.wf_delete_orders;
end if;

create TRIGGER wf_delete_orders before delete on
Orders
referencing old as old_name
for each row
begin
	declare remoteServer varchar(32);
	select sysname into remoteServer from guideventure where ventureId = old_name.ventureId;
	if remoteServer is not null then
		call delete_remote(remoteServer, 'jscet', 'id = ' + convert(varchar(20), old_name.id_jscet));
	end if;
--  delete from inv where id = old_name.id_inv;
end;


if exists (select '*' from sysprocedure where proc_name like 'wf_next_numdoc') then  
	drop procedure wf_next_numdoc;
end if;

create 
	function wf_next_numdoc() returns integer
begin
	declare sys_numdoc integer;
	declare sys_numdoc_c varchar(10);
	declare sys_year_i integer;
	declare sys_mmdd char(4);
	declare sys_number_c varchar(4);
--	declare sys_number_i integer;

	declare now_year_ln integer;
	declare now_date char(6);
	declare now_year_i integer;
	declare now_year char(2);
	declare now_mmdd char(4);
	declare now_m char(1);
	declare v_new_base integer;


	-- по умолчанию в том же дне
	set v_new_base = 0;

	select lastDocNum into sys_numdoc from system;
	set sys_numdoc_c = convert(varchar(10), sys_numdoc);

	set now_date = convert(char(6), now(), 12); -- 050716
	set now_year = substring(now_date, 1, 2);
	set now_year_i = convert(integer, now_year); --5 или 10 если 2010-й год
	set now_year_ln = char_length(convert(char(2), now_year_i)); --1 или 2

	-- Стандарная маска номера YMMDDnn[n..] 
	 
	set sys_year_i = convert(integer, substring(sys_numdoc_c, 1, now_year_ln));
	if (sys_year_i != now_year_i) then
		-- Переход на новый год
		set v_new_base = 1;
		-- Учесть переход с 31.12.2009 на 01.01.2010
		-- изменяется длина шаблона номера счета
		--if sys_year_i = 9 and now_year = 10 then
			--??? set v_year_now = 0;
		--end if;
	end if;

	
	set sys_mmdd = substring (sys_numdoc, now_year_ln + 1, 4);
	set now_m = convert(char(1), 2+convert(integer, convert(char(1), substring(now_date,3,1))));
	set now_mmdd = now_m + substring(now_date, 4, 3);
	if sys_mmdd != now_mmdd then
		set v_new_base = 1;
	end if;

	if v_new_base = 0 then
		set sys_number_c = substring (sys_numdoc_c, now_year_ln + 5);
		set sys_number_c = convert(varchar(3), convert(integer, sys_number_c) + 1);
		if char_length(sys_number_c) = 1 then
			set sys_number_c = '0' + sys_number_c;
		end if;
		set wf_next_numdoc = convert(char(2),sys_year_i) + sys_mmdd + sys_number_c;
	else 
		set wf_next_numdoc = convert(char(2),now_year_i) + now_mmdd + '01';
	end if;

	update system set lastDocNum = wf_next_numdoc;


end;


if exists (select '*' from sysprocedure where proc_name like 'wf_next_numorder') then  
	drop procedure wf_next_numorder;
end if;

create 
	function wf_next_numorder() returns integer
begin
	declare sys_numorder integer;
	declare sys_numorder_c varchar(10);
	declare sys_year_i integer;
	declare sys_mmdd char(4);
	declare sys_number_c varchar(4);
--	declare sys_number_i integer;

	declare now_year_ln integer;
	declare now_date char(6);
	declare now_year_i integer;
	declare now_year char(2);
	declare now_mmdd char(4);
	declare v_new_base integer;


	-- по умолчанию в том же дне
	set v_new_base = 0;

	select lastPrivatNum into sys_numorder from system;
	set sys_numorder_c = convert(varchar(10), sys_numorder);

	set now_date = convert(char(6), now(), 12); -- 050716
	set now_year = substring(now_date, 1, 2);
	set now_year_i = convert(integer, now_year); --5 или 10 если 2010-й год
	set now_year_ln = char_length(convert(char(2), now_year_i)); --1 или 2

	-- Стандарная маска номера YMMDDnn[n..] 
	 
	set sys_year_i = convert(integer, substring(sys_numorder_c, 1, now_year_ln));
	if (sys_year_i != now_year_i) then
		-- Переход на новый год
		set v_new_base = 1;
		-- Учесть переход с 31.12.2009 на 01.01.2010
		-- изменяется длина шаблона номера счета
		--if sys_year_i = 9 and now_year = 10 then
			--??? set v_year_now = 0;
		--end if;
	end if;

	
	set sys_mmdd = substring (sys_numorder, now_year_ln + 1, 4);
	set now_mmdd = substring (now_date, 3, 4);
	if sys_mmdd != now_mmdd then
		set v_new_base = 1;
	end if;

	if v_new_base = 0 then
		set sys_number_c = substring (sys_numorder_c, now_year_ln + 5);
		set sys_number_c = convert(varchar(3), convert(integer, sys_number_c) + 1);
		if char_length(sys_number_c) = 1 then
			set sys_number_c = '0' + sys_number_c;
		end if;
		set wf_next_numorder = convert(char(2),sys_year_i) + sys_mmdd + sys_number_c;
	else 
		set wf_next_numorder = convert(char(2),now_year_i) + now_mmdd + '01';
	end if;

	update system set lastPrivatNum = wf_next_numorder;

end;


-----------------------------------------------------
--	Функции, для работы с вариантными изделиями -----
-----------------------------------------------------

if exists (select 1 from sysprocedure where proc_name = 'wf_get_variant_Id') then
	drop procedure wf_get_variant_Id;
end if;


CREATE FUNCTION wf_get_variant_Id(
	p_numOrder varchar(50)
	, p_productid integer
	, p_prext integer
	, p_incompleteNomnom varchar(20) default null
)
returns integer
begin
	declare v_variantId integer;
	declare is_ok integer;

	-- курсор пробегает по всем комплектам вариантного изделия
	-- которые раньше уже были созданы
	declare c_product_variants dynamic scroll cursor for
		select id_variant from sguidecomplect
		where productId = p_productId;
	open c_product_variants;
	set is_ok = null;
	set v_variantId = 0;

	all_variants: loop
		fetch c_product_variants into v_variantId;
		if SQLCODE <>0 then 
			leave all_variants;
		end if;
		
		set is_ok = wf_try_variant(v_variantId, p_numOrder, p_productId, p_prExt, p_incompleteNomnom);
		if is_ok is not null then
			leave all_variants;
		end if;
	end loop;
	close c_product_variants;
	if is_ok is null then
		set v_variantId = wf_put_variant(p_numOrder, p_productId, p_prExt, p_incompleteNomnom);
	end if;
	return v_variantId;
end;


if exists (select 1 from sysprocedure where proc_name = 'wf_put_variant') then
	drop procedure wf_put_variant;
end if;

CREATE FUNCTION wf_put_variant(p_numOrder varchar(50), p_productid integer, p_prext integer, p_incompleteNomnom varchar(20) default null)
returns integer
begin
	declare order_nom char(50);
	declare v_variantId integer;
//	declare g_id integer; // Глобальный идентификатор на все сервера 
	declare v_xprext integer;
	declare v_nomNom varchar(30);
	declare v_nomName varchar(100);
	declare v_id_size integer;
	declare v_id_edizm integer;
	declare v_prc1 double;
	declare v_id_compl integer;
	declare v_id_Inv integer;
	declare v_id_Inv_compl integer;
	declare v_kol integer;
	declare v_belong_Id integer;
	declare v_variant_id integer;
	declare v_nm varchar(102);
	declare v_size varchar(30);

	declare v_table_name varchar(100);
	declare v_fields varchar(1000);
	declare v_values varchar(1000);

	declare c_order_nom dynamic scroll cursor for
		select nomNom 
		from xVariantnomenc vn
		where vn.prId = p_productId and vn.prExt = p_prExt and vn.numOrder = p_numOrder

				union

		select nomNom from sproducts p
		where 
			    p.productId = p_productId
			and exists (select 1 from svariantpower vp where vp.productid = p.productid)
			and not exists (select 1 from sguidevariant gv where p.productid = gv.productid and p.xgroup = gv.xgroup and not (gv.xgroup = '' or gv.c = 1))

				union

	    select p_incompleteNomnom 
	    where p_incompleteNomnom is not null
		order by 1;

	select max(xPrExt)into v_xprExt from sguideComplect where productId = p_productId; 
	set v_xPrExt = isnull(v_xPrExt, 0) + 1 ;

	// Здесь нужно вставить добавление во все slave.inv таблицы новый комплект вариантного изделия
	//  v_id_inv - новый вариант вариантного изделия
	//  v_belong_id - id папки, которая объединяет все варианты вариантного издлия
	// -----------------------
	set v_id_inv = get_nextid('inv');

		select 
			  prName as v_nomNom
			, prDescript as v_nomName
			, prSize as v_size
			, s.id_size
			, e1.id_edizm as v_id_edizm
			, n.cena4 as v_prc1
			, n.id_inv as v_belong_id
		into
			  v_nomNom
			, v_nomName
			, v_size
			, v_id_size
			, v_id_edizm
			, v_prc1
			, v_belong_id
		from sguideproducts n
		join sguideseries p on p.seriaid = n.prseriaid
		join edizm e1 on e1.name = 'шт.'
		left join size s on s.name = n.prsize
		where n.prid = p_productid;

		set v_id_size = isnull(v_id_size, 0);
		
		// теперь это изделие обязано быть группой,
		// под которой уже будут собираться все варианты
		call update_host('inv', 'is_group', '1', 'id = ' + convert(varchar(20), v_id_inv));

		set v_nm = wf_make_variant_nm (
			  v_nomName
			, v_size
			, v_nomNom
			, v_xprext
		);
	
		// Добавляем вариант в подгруппу		
		set v_fields ='id'
		+ ', belong_id'
		+ ', nomen'
		+ ', nm'
		+ ', id_edizm1'
		+ ', id_size'
		+ ', prc1'
		+ ', is_compl'
		;
		set v_values =
			 convert(varchar(20), v_id_inv)
			+ ', ' + convert(varchar(20), v_belong_id)
			+ ', ''''' + v_nomnom + ''''''
			+ ', ''''' + v_nm + ''''''
			+ ', ' + convert(varchar(20), v_id_edizm)
			+ ', ' + convert(varchar(20), v_id_size)
			+ ', ''''' + convert(varchar(20), v_prc1) + ''''''
			+ ', 1'
		;	
    
		call insert_host ('inv', v_fields, v_values);
	
	
	// Заглолвок комплекта
	insert into sguidecomplect (productId, xPrExt, id_inv)
		values (p_productId, v_xPrExt, v_id_inv);
		
	set v_variantId = @@identity;
		
	open c_order_nom;
	find: loop
		fetch c_order_nom into order_nom;
		if SQLCODE != 0 then
			leave find;
		end if;
		
		// А здесь в slave.compl
		//  ...
		//
		set v_id_compl = get_nextid('compl');
		
		select n.id_inv 
			, e.id_edizm
			, p.quantity
		into 
			v_id_inv_compl
			, v_id_edizm
			, v_kol
		from sproducts p 
		join sguidenomenk n on n.nomnom = order_nom and p.nomnom = n.nomnom
		join edizm e on e.name = n.ed_izmer
		where p.productid = p_productid;

		
		// Добавляем комплектацию варианта во все бфзы
		set v_fields ='id'
		+ ', id_inv'
		+ ', id_inv_belong'
		+ ', id_edizm'
		+ ', kol'
		;
		set v_values =
			 convert(varchar(20), v_id_compl)
			+ ', ' + convert(varchar(20), v_id_inv_compl)
			+ ', ' + convert(varchar(20), v_id_inv)
			+ ', ' + convert(varchar(20), v_id_edizm)
			+ ', ''''' + convert(varchar(20), v_kol) + ''''''
		;	
    
		call insert_host ('compl', v_fields, v_values);

				
		insert into svariantcomplect (id_variant, nomnom, id_compl)
		values (v_variantId, order_nom, v_id_compl);
	end loop;
	close c_order_nom;
	
	return v_variantId;
	
end;



if exists (select 1 from sysprocedure where proc_name = 'wf_try_variant') then
	drop function wf_try_variant;
end if;

CREATE FUNCTION "wf_try_variant"(p_id_variant integer, p_numOrder varchar(50), p_productid integer, p_prext integer, p_incompleteNomnom varchar(20) default null) returns integer
begin
	
	declare variant_nom char(50);
	declare order_nom char(50);
	declare is_variant_end integer;
	declare is_order_end integer;
	declare ret integer;
	
	declare c_order_nom dynamic scroll cursor for
		select nomNom 
		from xVariantnomenc vn
		where vn.prId = p_productid and vn.prExt = p_prExt and vn.numOrder = p_numorder
				union
		select nomNom from sproducts p
		where 
			    p.productId = p_productId
			and exists (select 1 from svariantpower vp where vp.productid = p.productid)
			and not exists (select 1 from sguidevariant gv where p.productid = gv.productid and p.xgroup = gv.xgroup)
--			and exists (select 1 from xVariantNomenc vn where vn.prId = p.productId and vn.prId = p_productid and vn.prExt = p_prExt and vn.numOrder = p_numorder)
				union
	    select p_incompleteNomnom 
	    where p_incompleteNomnom is not null
	    order by 1;

	declare c_variant_nom dynamic scroll cursor for
		select nomnom from svariantcomplect vc
		where vc.id_variant = p_id_variant
		order by 1;

	open c_order_nom;
	open c_variant_nom;
	set ret = null;
	find: loop
		set is_order_end = 0;
		fetch c_order_nom into order_nom;
		if SQLCODE != 0 then
			set is_order_end = 1;
		end if;
		set is_variant_end = 0;
		fetch c_variant_nom into variant_nom;
		if SQLCODE != 0 then
			set is_variant_end = 1;
		end if;
		if is_order_end = 1 and is_variant_end = 1 then
			set ret = 1; -- success!
			leave find;
		end if;
		if variant_nom is null or order_nom is null then
			leave find;
		end if;
		if is_order_end = 1 or is_variant_end = 1 or variant_nom != order_nom then
			leave find;
		end if;
	end loop;
	close c_variant_nom;
	close c_order_nom;
	return ret;
end;

if exists (select 1 from sysprocedure where proc_name = 'get_currency_rate_id') then
	drop function get_currency_rate_id;
end if;

if exists (select 1 from sysprocedure where proc_name = 'system_currency') then
	drop function system_currency;
end if;

create function system_currency(
	)
	returns integer
begin
	select id_cur into system_currency from system;
end;


if exists (select 1 from sysprocedure where proc_name = 'system_currency_rate') then
	drop function system_currency_rate;
end if;

create function system_currency_rate(
	)
	returns float
begin
	select abs(kurs) into system_currency_rate from system;
end;




----------------------------------------------------------------------
--------------         xPredmetyByNomenkOut          -----------------
----------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_xPredmetyByNomenkOut_outcome_di' and tname = 'xPredmetyByNomenkOut') then 
	drop trigger xPredmetyByNomenkOut.wf_xPredmetyByNomenkOut_outcome_di;
end if;
/*
create 
	trigger wf_xPredmetyByNomenkOut_outcome_di before delete order 1 on 
xPredmetyByNomenkOut
referencing old as old_name
for each row
begin
--	declare v_id_mat integer;
	declare v_id_jmat integer;
	declare v_sysname varchar(50);

--	set v_id_mat = old_name.id_mat;
	set v_id_jmat = old_name.id_jmat;

	select v.sysname
	into v_sysname
	from orders o
		left join guideventure v on v.ventureid = o.ventureid and v.standalone = 0
	where numorder = old_name.numorder;

	call wf_otgruz_remove (
		v_id_jmat
		,'stime'
	);

	if v_sysname is not null and v_sysname != 'stime' then
		call wf_otgruz_remove (
			v_id_jmat
			,v_sysname
		);

	end if;

		
		
end;
*/




if exists (select 1 from systriggers where trigname = 'wf_xPredmetyByNomenkOut_outcome_ui' and tname = 'xPredmetyByNomenkOut') then 
	drop trigger xPredmetyByNomenkOut.wf_xPredmetyByNomenkOut_outcome_ui;
end if;
/*
create 
	trigger wf_xPredmetyByNomenkOut_outcome_ui before update order 1 on 
xPredmetyByNomenkOut
referencing new as new_name old as old_name
for each row
begin
	declare v_id_mat integer;
	declare v_id_jmat integer;
	declare v_sysname varchar(50);
	declare v_cena float;

	if update(quant) and old_name.quant != new_name.quant then
		set v_id_mat = old_name.id_mat;
		set v_id_jmat = old_name.id_jmat;

		select cenaEd into v_cena from xPredmetybyNomenk where numOrder = new_name.numOrder and nomnom = new_name.nomNom;

		select v.sysname
		into v_sysname
		from orders o
		left join guideventure v on v.ventureid = o.ventureid and v.standalone = 0
		where numorder = old_name.numorder;
		

		call wf_otgruz_quant(
			v_id_mat
			,v_id_jmat
			,new_name.quant
			,v_cena
			,'stime'
		);

		if v_sysname is not null and v_sysname != 'stime' then
			call wf_otgruz_quant(
				v_id_mat
				,v_id_jmat
				,new_name.quant
				,v_cena
				,v_sysname
			);

		end if;


	end if;
end;
*/

if exists (select 1 from systriggers where trigname = 'wf_xPredmetyByNomenkOut_outcome_bi' and tname = 'xPredmetyByNomenkOut') then 
	drop trigger xPredmetyByNomenkOut.wf_xPredmetyByNomenkOut_outcome_bi;
end if;

/*
create 
	trigger wf_xPredmetyByNomenkOut_outcome_bi before insert order 1 on 
xPredmetyByNomenkOut
referencing new as new_name
for each row
begin
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_jmat_nu varchar(20);
	declare v_mat_nu varchar(20);
	declare v_currency_rate float;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_id_source integer;
	declare v_id_dest integer;
--	declare v_osn varchar(100);
	declare v_id_jscet integer;
--	declare v_venture_id integer;
	declare v_firm_id integer;
	declare v_sysname varchar(50);
	declare v_ventureName varchar(100);
	declare v_cena float;
	declare v_cur_otgruz_date date;


	if get_standalone('stime') = 1 then
		call log_warning('Информация об отгрузке по заказу ' + convert(varchar(20), new_name.numorder) + ' не попадает в аналитическую базу stime.');
		return;
	end if;

	select max(id_jmat) into v_id_jmat 
	from xPredmetyByIzdeliaOut 
	where numOrder = new_name.numorder and outDate = new_name.outDate;

	if v_id_jmat is null then
		select max(id_jmat) into v_id_jmat 
		from xPredmetyByNomenkOut 
		where numOrder = new_name.numorder and outDate = new_name.outDate;
	end if;

	select 
		 o.id_jscet
		, isnull(s.id_voc_names, 0)
		, isnull(f.id_voc_names,0)
		, v.ventureName
		, v.sysname
	into  
		 v_id_jscet
		, v_id_source
		, v_id_dest
		, v_ventureName
		, v_sysname
	from orders o
		left join guideventure v on v.ventureid = o.ventureid and v.standalone = 0
		left join guidefirms f on o.firmid = f.firmid
		left join sguidesource s on sourceid = -1001
	where numorder = new_name.numorder;

	
	set v_id_currency = system_currency();
	call slave_currency_rate_stime(v_datev, v_currency_rate);

--	select id_voc_names into v_id_dest from guidefirms where firmid = v_firm_id;
--	    message 'v_id_dest = ', v_id_dest to client;
	-- со склада 1 
	-- ?? хотя по идее нужно бы отгружать со склада готовой продукции
--	select id_voc_names into v_id_source from sguidesource where sourceid = -1001;

	if v_id_jmat is null then
--	    message '---' to client;
		set v_id_jmat = wf_otgruz_jmat(
			new_name.numorder
			, v_id_jscet
--			, v_venture_id
			, new_name.outDate
			, v_id_source
			, v_id_dest
			, v_id_currency
			, v_datev
			, v_currency_rate
			, v_sysname
		);
--		update orders set id_jmat = v_id_jmat where numorder = new_name.numorder;
	end if;

	set v_id_mat = get_nextid('mat');
	call slave_select_stime(v_mat_nu, 'mat', 'max(nu)', 'id_jmat = ' + convert(varchar(20), v_id_jmat));
	set v_mat_nu = convert(varchar(20), convert(integer, isnull(v_mat_nu, 0)) + 1);
	select cenaEd into v_cena from xPredmetybyNomenk where numOrder = new_name.numOrder and nomnom = new_name.nomNom;

	call wf_otgruz_nom(
		  v_id_mat
		, v_id_jmat
		, new_name.nomnom
		, new_name.quant
		, v_cena
		, v_mat_nu
		, v_id_source
		, v_id_dest
		, v_currency_rate
		, v_sysname
	);
	set new_name.id_mat = v_id_mat;
	set new_name.id_jmat = v_id_jmat;

end;
*/


----------------------------------------------------------------------
--------------       Otgruz helpers PROCEDURIES           ------------
----------------------------------------------------------------------
if exists (select 1 from sysprocedure where proc_name = 'wf_otgruz_remove') then
	drop procedure wf_otgruz_remove;
end if;

/*
CREATE procedure wf_otgruz_remove(
	  p_id_jmat integer
	, p_sysname varchar(50)
) 
begin
	execute immediate 'call slave_delete_'+p_sysname+'(''jmat''
		, ''id = '' + convert(varchar(20), '+convert(varchar(20), p_id_jmat)+'))'
	;
	execute immediate 'call slave_delete_'+p_sysname+'(''mat''
		, ''id_jmat = '' + convert(varchar(20), '+convert(varchar(20), p_id_jmat)+'))'
	;
end;
*/


if exists (select 1 from sysprocedure where proc_name = 'wf_otgruz_quant') then
	drop procedure wf_otgruz_quant;
end if;

/*
CREATE procedure wf_otgruz_quant(
	  p_id_mat integer
	, p_id_jmat integer
	, p_quant  float
	, p_cena float
--	, p_currency_rate float
	, p_sysname varchar(50)
) 
begin
	declare v_currency_rate float;


	execute immediate 'call slave_select_'+p_sysname+'(v_currency_rate, ''jmat'', ''curr'', ''id = '' + convert(varchar(20), ' + convert(varchar(20), p_id_jmat) +'))';
--	select v_currency_rate;
	execute immediate 'call slave_update_'+p_sysname+'(''mat''
		, ''kol1''
		, '''+convert(varchar(20), round(p_quant, 2))+'''
		, ''id = '' + convert(varchar(20), '+convert(varchar(20), p_id_mat)+'))';
	execute immediate 'call slave_update_'+p_sysname+'(''mat''
		, ''kol3''
		, '''+convert(varchar(20), round(p_quant, 2))+'''
		, ''id = '' + convert(varchar(20), '+convert(varchar(20), p_id_mat)+'))';
	execute immediate 'call slave_update_'+p_sysname+'(''mat''
		, ''summa_sale''
		, '''+convert(varchar(20), round(p_quant* p_cena * v_currency_rate, 2))+'''
		, ''id = '' + convert(varchar(20), '+convert(varchar(20), p_id_mat)+'))';
	execute immediate 'call slave_update_'+p_sysname+'(''mat''
		, ''summa_salev''
		, '''+convert(varchar(20), round(p_quant* p_cena, 2))+'''
		, ''id = '' + convert(varchar(20), '+convert(varchar(20), p_id_mat)+'))';
end;
*/

if exists (select 1 from sysprocedure where proc_name = 'wf_otgruz_nom') then
	drop procedure wf_otgruz_nom;
end if;
/*

CREATE procedure wf_otgruz_nom(
	  p_id_mat integer
	, p_id_jmat integer
	, p_nomnom varchar(50)
	, p_quant  float
	, p_cena float
	, p_mat_nu varchar(20)
	, p_id_source integer
	, p_id_dest integer
	, p_currency_rate float
	, p_sysname varchar(50) default null
) 
begin
--	declare v_id_jmat integer;
--	declare v_id_mat integer;
--	declare v_mat_nu varchar(20);
--	declare v_currency_rate float;
--	declare v_datev date;
--	declare v_id_currency integer;
	declare v_id_inv integer;
--	declare v_id_source integer;
--	declare v_id_dest integer;
--	declare v_cost float;
	declare v_perList float;

	declare sync char(1);

	select id_inv, perList into v_id_inv, v_perList from sguidenomenk where nomnom = p_nomnom;


--	call call_host('block_table', 'sync, ''prior'', ''mat''');
	call block_remote('stime', @@servername, 'mat');
	
	call wf_insert_mat (
		'stime'
		,p_id_mat
		,p_Id_jmat
		,v_id_inv
		,p_mat_nu
		,p_quant 
		,p_cena
		,p_currency_rate
		,p_id_source
		,p_id_dest
		,v_perList
	);
	
	if p_sysname is not null and p_sysname != 'stime' then
		call wf_insert_mat (
			p_sysname
			,p_id_mat
			,p_Id_jmat
			,v_id_inv
			,p_mat_nu
			,p_quant 
			,p_cena
			,p_currency_rate
			,p_id_source
			,p_id_dest
			,v_perList
		);
	
	end if;


	call block_remote('stime', @@servername, 'mat');
--	call call_host('unblock_table', 'sync, ''prior'', ''mat''');
	--	set wf_otgruz_nom = v_id_mat;
end;
*/


	
	
if exists (select 1 from sysprocedure where proc_name = 'wf_otgruz_jmat') then
	drop function wf_otgruz_jmat;
end if;

/*
CREATE FUNCTION wf_otgruz_jmat(
	p_numorder integer
	,p_id_jscet integer
--	,p_venture_id integer
	,p_date date
	,p_id_source integer
	,p_id_dest integer
	,p_id_currency integer
	,p_datev date
	,p_currency_rate float
	,p_sysname varchar(50) default null
) 
	returns integer
begin
	
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_jmat_nu varchar(20);
--	declare v_currency_rate float;
--	declare v_datev date;
	declare v_id_currency integer;
--	declare v_id_source integer;
--	declare v_id_dest integer;
	declare v_osn varchar(100);
--	declare v_sysname varchar(50);
--	declare v_ventureName varchar(200);

		set v_id_jmat = get_nextid('jmat');
--		set v_id_currency = system_currency();
--		call slave_currency_rate_stime(v_datev, v_currency_rate);
		set v_jmat_nu = nextnu_remote('stime', 'jmat');
		--select id_voc_names into v_id_source from sguidesource where sourceid = new_name.sourid;
		--select id_voc_names into v_id_dest from sguidesource where sourceid = new_name.destid;

		set v_osn = 'заказ N ' + convert(varchar(20), p_numorder);
	    
		call wf_insert_jmat (
			'stime'
			,'1210' --v_id_guide_jmat
			,v_id_jmat
			,p_date --v_jmat_date
			,v_jmat_nu
			,v_osn
			,p_id_currency
			,p_datev
			,p_currency_rate
			,p_id_source
			,p_id_dest
		);

		if p_sysname is not null and p_sysname != 'stime' then
			call wf_insert_jmat (
				p_sysname
				,'1210' --v_id_guide_jmat
				,v_id_jmat
				,p_date --v_jmat_date
				,v_jmat_nu
				,v_osn
				,p_id_currency
				,p_datev
				,p_currency_rate
				,p_id_source
				,p_id_dest
				,p_id_jscet
			);
		end if;

		set wf_otgruz_jmat = v_id_jmat;

end;
*/


----------------------------------------------------------------------
--------------         xPredmetyByIzdeliaOut          ----------------
----------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_xPredmetyByIzdeliaOut_outcome_di' and tname = 'xPredmetyByIzdeliaOut') then 
	drop trigger xPredmetyByIzdeliaOut.wf_xPredmetyByIzdeliaOut_outcome_di;
end if;

/*
create 
	trigger wf_xPredmetyByIzdeliaOut_outcome_di before delete order 1 on 
xPredmetyByIzdeliaOut
referencing old as old_name
for each row
begin

--	declare v_id_mat integer;
	declare v_id_jmat integer;
	declare v_sysname varchar(50);

--	set v_id_mat = old_name.id_mat;
	set v_id_jmat = old_name.id_jmat;

	select v.sysname
	into v_sysname
	from orders o
		left join guideventure v on v.ventureid = o.ventureid and v.standalone = 0
	where numorder = old_name.numorder;

	call wf_otgruz_remove (
		v_id_jmat
		,'stime'
	);

	if v_sysname is not null and v_sysname != 'stime' then
		call wf_otgruz_remove (
			v_id_jmat
			,v_sysname
		);

	end if;

		
		
end;

*/



if exists (select 1 from systriggers where trigname = 'wf_xPredmetyByIzdeliaOut_outcome_ui' and tname = 'xPredmetyByIzdeliaOut') then 
	drop trigger xPredmetyByIzdeliaOut.wf_xPredmetyByIzdeliaOut_outcome_ui;
end if;

/*
create 
	trigger wf_xPredmetyByIzdeliaOut_outcome_ui before update order 1 on 
xPredmetyByIzdeliaOut
referencing new as new_name old as old_name
for each row
begin

	declare v_id_mat integer;
	declare v_id_jmat integer;
	declare v_sysname varchar(50);
	declare v_cena float;

	if update(quant) and old_name.quant != new_name.quant then
		set v_id_mat = old_name.id_mat;
		set v_id_jmat = old_name.id_jmat;

		select cenaEd into v_cena 
		from xPredmetybyIzdelia pi
		where numOrder = new_name.numOrder 
			and pi.prId = new_name.prId 
			and pi.prExt = new_name.prExt
		;

		select v.sysname
		into v_sysname
		from orders o
		left join guideventure v on v.ventureid = o.ventureid and v.standalone = 0
		where numorder = old_name.numorder;
		

		call wf_otgruz_quant(
			v_id_mat
			,v_id_jmat
			,new_name.quant
			,v_cena
			,'stime'
		);

		if v_sysname is not null and v_sysname != 'stime' then
			call wf_otgruz_quant(
				v_id_mat
				,v_id_jmat
				,new_name.quant
				,v_cena
				,v_sysname
			);

		end if;


	end if;

end;
*/

if exists (select 1 from systriggers where trigname = 'wf_xPredmetyByIzdeliaOut_outcome_bi' and tname = 'xPredmetyByIzdeliaOut') then 
	drop trigger xPredmetyByIzdeliaOut.wf_xPredmetyByIzdeliaOut_outcome_bi;
end if;
/*
create 
	trigger wf_xPredmetyByIzdeliaOut_outcome_bi before insert order 1 on 
xPredmetyByIzdeliaOut
referencing new as new_name
for each row
begin
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_jmat_nu varchar(20);
--	declare v_mat_nu varchar(20);
	declare v_currency_rate float;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_id_source integer;
	declare v_id_dest integer;
--	declare v_osn varchar(100);
	declare v_id_jscet integer;
--	declare v_venture_id integer;
	declare v_firm_id integer;
	declare v_sysname varchar(50);
	declare v_ventureName varchar(100);
	declare v_cena float;
	declare v_cur_otgruz_date date;



--	set v_id_jmat = old_name.id_jmat;
	select max(id_jmat) into v_id_jmat 
	from xPredmetyByIzdeliaOut 
	where numOrder = new_name.numorder and outDate = new_name.outDate;

	if v_id_jmat is null then
		select max(id_jmat) into v_id_jmat 
		from xPredmetyByNomenkOut 
		where numOrder = new_name.numorder and outDate = new_name.outDate;
	end if;

	select 
		 o.id_jscet
		, isnull(s.id_voc_names, 0)
		, isnull(f.id_voc_names,0)
		, v.ventureName
		, v.sysname
	into  
		 v_id_jscet
		, v_id_source
		, v_id_dest
		, v_ventureName
		, v_sysname
	from orders o
		left join guideventure v on v.ventureid = o.ventureid and v.standalone = 0
		left join guidefirms f on o.firmid = f.firmid
		left join sguidesource s on sourceid = -1001
	where numorder = new_name.numorder;

	
	set v_id_currency = system_currency();
	call slave_currency_rate_stime(v_datev, v_currency_rate);

	if v_id_jmat is null then
		set v_id_jmat = wf_otgruz_jmat(
			new_name.numorder
			, v_id_jscet
--			, v_venture_id
			, new_name.outDate
			, v_id_source
			, v_id_dest
			, v_id_currency
			, v_datev
			, v_currency_rate
			, v_sysname
		);
	end if;

--	message 'v_id_jscet = ', v_id_jscet to client;
--	message 'v_id_jmat = ', v_id_jmat to client;
	set v_id_mat = get_nextid('mat');
--	call slave_select_stime(v_mat_nu, 'mat', 'max(nu)', 'id_jmat = ' + convert(varchar(20), v_id_jmat));
--	set v_mat_nu = convert(varchar(20), convert(integer, isnull(v_mat_nu, 0)) + 1);

	select cenaEd 
	into v_cena
	from xPredmetybyIzdelia pi
	where pi.numOrder = new_name.numOrder 
		and pi.prId = new_name.prId 
		and pi.prExt = new_name.prExt;

	call wf_otgruz_izd(
		  v_id_mat
		, v_id_jmat
		, new_name.numOrder
		, new_name.prId
		, new_name.prExt
		, new_name.quant
		, v_cena
--		, v_mat_nu
		, v_id_source
		, v_id_dest
		, v_currency_rate
		, v_sysname
	);
	
	
	set new_name.id_mat = v_id_mat;
	set new_name.id_jmat = v_id_jmat;

end;
*/

if exists (select 1 from sysprocedure where proc_name = 'wf_otgruz_izd') then
	drop procedure wf_otgruz_izd;
end if;
/*

CREATE procedure wf_otgruz_izd(
	  p_id_mat integer
	, p_id_jmat integer
	, p_numOrder integer
	, p_prId integer
	, p_prExt integer
	, p_quant  float
	, p_cena float
--	, p_mat_nu varchar(20)
	, p_id_source integer
	, p_id_dest integer
	, p_currency_rate float
	, p_sysname varchar(50) default null
) 
begin
--	declare v_id_jmat integer;
--	declare v_id_mat integer;
	declare v_mat_nu varchar(20);
--	declare v_currency_rate float;
--	declare v_datev date;
--	declare v_id_currency integer;
	declare v_id_inv integer;
--	declare v_id_source integer;
--	declare v_id_dest integer;
--	declare v_cost float;
--	declare v_quant float;

	declare sync char(1);

	
--	call call_host('block_table', 'sync, ''prior'', ''mat''');

	call slave_select_stime(v_mat_nu, 'mat', 'max(nu)', 'id_jmat = ' + convert(varchar(20), p_id_jmat));
--	set v_mat_nu = convert(varchar(20), convert(integer, isnull(v_mat_nu, 0)) + 1);
   
	for aCursor as a dynamic scroll cursor for
	
		select id_inv r_id_inv, cost as r_cost, pr.quantity r_quant, nom.perList as r_perList
		from sguidenomenk nom
		join (
			select nomnom, quantity 
			from sproducts p
			where p.productId = p_prId 
			and ( exists (
					select 1 from sguidevariant vp 
					where 
						p.productid = vp.productid and p.xgroup = vp.xgroup and vp.c = 1
					)
					or p.xgroup = '' or p.xgroup is null
				)
					union 
			select v.nomnom, p.quantity 
			from xvariantnomenc v
				join sproducts p on p.productid = v.prid and p.nomnom = v.nomnom
			where v.prid = p_prId and v.numorder = p_numOrder and v.prExt=p_prExt
		) pr on pr.nomnom = nom.nomnom
	do
	
		
		set v_mat_nu = convert(varchar(20), convert(integer, isnull(v_mat_nu, 0)) + 1);
		call wf_insert_mat (
			'stime'
			,p_id_mat
			,p_Id_jmat
			,r_id_inv
			,v_mat_nu
			,p_quant * r_quant
			,r_cost
			,p_currency_rate
			,p_id_source
			,p_id_dest
			,r_perList
		);
		set p_id_mat = p_id_mat + 1;

	end for;

	
	if p_sysname is not null and p_sysname != 'stime' then
	
--	message 'p_prId = ', p_prId to client;
		if not exists (select 1 from svariantpower where productId = p_prid) then
			select id_inv into v_id_inv from sguideproducts where prId = p_prId;
		else 
			select id_inv into v_Id_inv 
			from xPredmetyByIzdelia pi 
			where pi.prId = p_prId 
				and pi.prExt = p_prExt
				and pi.numOrder = p_numOrder
		end if;
--	message 'v_id_inv = ', v_id_inv to client;

		execute immediate 'call slave_select_'+p_sysname+'(v_mat_nu, ''mat'', ''max(nu)'''
			+', ''id_jmat = '' + convert(varchar(20), '+convert(varchar(20), p_id_jmat)+'))';

		set v_mat_nu = convert(varchar(20), convert(integer, isnull(v_mat_nu, 0)) + 1);

		call wf_insert_mat (
			p_sysname
			,p_id_mat
			,p_Id_jmat
			,v_id_inv
			,v_mat_nu
			,p_quant 
			,p_cena
			,p_currency_rate
			,p_id_source
			,p_id_dest
			,1
		);
	
	end if;


--	call call_host('unblock_table', 'sync, ''prior'', ''mat''');
	--	set wf_otgruz_izd = v_id_mat;
end;
*/
	



-------------------------------------------------------------------------
--------------             BayOrders      ----------------------------
-------------------------------------------------------------------------

if exists (select 1 from systriggers where trigname = 'wf_insert_orders' and tname = 'BayOrders') then 
	drop trigger BayOrders.wf_insert_orders;
end if;

create TRIGGER "wf_insert_orders" before insert on
BayOrders
referencing new as new_name
for each row
begin
end;


if exists (select 1 from systriggers where trigname = 'wf_update_orders' and tname = 'BayOrders') then 
	drop trigger BayOrders.wf_update_orders;
end if;

create TRIGGER wf_update_orders before update on
BayOrders
referencing old as old_name new as new_name
for each row
begin
	declare remoteServerOld varchar(32);
	declare remoteServerNew varchar(32);
	declare v_fields varchar(255);
	declare v_values varchar(2000);
	declare v_nu_jscet integer;
	declare r_nu varchar(50);
	declare r_id integer;
--	declare v_firm_id integer;
	declare v_invCode varchar(10);
	declare v_id_dest integer;
	declare v_id_schef integer;
	declare v_id_bux integer;
	declare v_id_bank integer;
	declare v_datev varchar(20);
	declare v_id_cur integer;
	declare v_currency_rate float;
	declare v_inv_date varchar(20);
	declare v_numOrder integer;

	declare sync char(1);

	if update(ventureId) then
		if new_name.ventureId = 0 then
			set new_name.ventureid = null;
		end if;
		if isnull(old_name.ventureId, 0) != isnull(new_name.ventureId, 0) then
			select sysname into remoteServerOld from GuideVenture where ventureId = old_name.ventureId;
			if remoteServerOld is not null then
				call delete_remote(remoteServerOld, 'jscet', 'id = ' + convert(varchar(20), old_name.id_jscet));
				call delete_remote(remoteServerOld, 'scet', 'id_jmat = ' + convert(varchar(20), old_name.id_jscet));
				set new_name.invoice = 'счет ?';
			end if;

			select sysname, invCode into remoteServerNew, v_invcode from GuideVenture where ventureId = new_name.ventureId;

			--message 'sysname = ', remoteServerNew to client;
			
			if remoteServerNew is not null then
	
				set v_numOrder = old_name.numOrder;
--				set v_firm_id = old_name.firmId;
				select id_voc_names into v_id_dest from bayguidefirms where firmid = old_name.firmId;
				call put_jscet(r_id, v_nu_jscet, remoteServerNew, v_numOrder, v_id_dest, old_name.invoice);
		
				set new_name.id_jscet = r_id;
				set new_name.invoice = v_invCode + convert(varchar(20), v_nu_jscet);
				call wf_set_bay_detail(remoteServerNew, r_id, new_name.numOrder, v_inv_date);
			end if;
		end if;
	end if;
	if update (firmId) then
		select sysname into remoteServerOld from GuideVenture where ventureId = old_name.ventureId;
		if remoteServerOld is not null then
			select id_voc_names into v_id_dest from BayGuideFirms where firmId = new_name.firmId;
			call block_remote(remoteServerOld, @@servername, 'jscet');
			call update_remote(remoteServerOld, 'jscet', 'id_d', convert(varchar(20), v_id_dest ), 'id = ' + convert(varchar(20), old_name.id_jscet));
			call update_remote(remoteServerOld, 'jscet', 'id_d_cargo', convert(varchar(20), v_id_dest ), 'id = ' + convert(varchar(20), old_name.id_jscet));
			call unblock_remote(remoteServerOld, @@servername, 'jscet');
		end if;
	end if;
end;


if exists (select 1 from systriggers where trigname = 'wf_delete_orders' and tname = 'BayOrders') then 
	drop trigger BayOrders.wf_delete_orders;
end if;

create TRIGGER wf_delete_orders before delete on
BayOrders
referencing old as old_name
for each row
begin
	declare remoteServer varchar(32);
	select sysname into remoteServer from guideventure where ventureId = old_name.ventureId;
	if remoteServer is not null then
		call delete_remote(remoteServer, 'jscet', 'id = ' + convert(varchar(20), old_name.id_jscet));
	end if;
--  delete from inv where id = old_name.id_inv;
end;



-- Процедура синхронизирует предметы bay-заказа Приора
-- с предметами счета в бухгалтерской базе комтеха
-- Это нужно сделать, если в заказ сначала 
-- добавть предметы, а только потом назначить предприятие,
-- через которую этот заказ должен пройти.
if exists (select '*' from sysprocedure where proc_name like 'wf_set_bay_detail') then  
	drop procedure wf_set_bay_detail;
end if;

create procedure wf_set_bay_detail (
			p_srvName varchar(20)
			, p_id_jscet integer
			, p_numOrder integer
			, p_date date
)
begin

	declare v_id_scet integer;
	declare v_id_inv integer;
	declare is_variant integer;
	declare v_id_variant integer;
	declare v_quant float;

	for c_nomenk as n dynamic scroll cursor for
		select 
			  p.nomNom as r_nomNom
			, p.quantity as r_quantity
			, intQuant as r_cenaEd
		from sDmcRez p
		where p.numDoc = p_numOrder
	do

		select 
			n.id_inv
			, r_quantity/n.perList
		into 
			v_id_inv
			, v_quant
		from 
			sGuideNomenk n
		where
			n.nomNom = r_nomNom;


		set v_id_scet = 
			wf_insert_scet(
				p_srvName
				, p_id_jscet
				, v_id_inv
				, v_quant
				, r_cenaEd
				, p_date
			);
		update sDmcRez set id_scet = v_id_scet where current of n;

	end for;



end;



-------------------------------------------------------------------------
-------------------             sDmcRez          ------------------------
-------------------------------------------------------------------------
--select * from scet_pm order by id_jmat desc
--select * from sDmcRez order by 1 desc
--select max(nu)+1  from scet_pm where id_jmat = 13281



if exists (select 1 from systriggers where trigname = 'wf_insert_nomenk' and tname = 'sDmcRez') then 
	drop trigger sDmcRez.wf_insert_nomenk;
end if;

create TRIGGER wf_insert_nomenk before insert on
sDmcRez
referencing new as new_name
for each row
begin
	declare v_id_scet integer;
	declare v_id_jscet integer;
	declare v_id_inv integer;
	declare remoteServerNew varchar(32);
	declare v_invCode varchar(10);
	declare v_date date;
	declare v_cenaEd float;
	declare v_quantity float;
	declare v_perList float;


--	message 'sDmcRez.wf_insert_nomenk' to client;
	select 
		o.id_jscet, o.inDate  
		, v.sysname, v.invCode
		, n.id_inv, n.perList 
	into 
		v_id_jscet, v_date 
		, remoteServerNew, v_invcode
		, v_id_inv, v_perList 
	from BayOrders o
	left join GuideVenture v on v.ventureid = o.ventureid and v.standalone = 0
	join sGuideNomenk n on n.nomNom = new_name.nomNom
	where o.numOrder = new_name.numDoc;


	set v_cenaEd = new_name.intQuant;
	set v_quantity = new_name.quantity/v_perList;

--	select id_inv into v_id_inv from sGuideNomenk where nomNom = new_name.nomNom;

--	select sysname, invCode into remoteServerNew, v_invcode from GuideVenture where ventureId = v_ventureId;

	if remoteServerNew is not null and v_id_jscet is not null then
	  -- Заказ, который имеет ссылки в бух.базах интеграции
	  -- т.е. уже назначен той, иди другой фирме
		set new_name.id_scet = 
			wf_insert_scet(
				remoteServerNew
				, v_id_jscet
				, v_id_inv
				, v_quantity
				, v_cenaEd
				, v_date
			);
	end if;
	  
end;


if exists (select 1 from systriggers where trigname = 'wf_update_nomenk' and tname = 'sDmcRez') then 
	drop trigger sDmcRez.wf_update_nomenk;
end if;

create TRIGGER "wf_update_nomenk" before update on
sDmcRez
referencing old as old_name new as new_name
for each row
begin
	declare v_id_scet integer;
	declare remoteServerNew varchar(32);

	declare v_cenaEd float;
	declare v_quantity float;
	declare v_perList float;
	
	set v_id_scet = old_name.id_scet;

	select v.sysname
		, n.perList 
	into remoteServerNew
		, v_perList 
	from BayOrders o
	join GuideVenture v on o.ventureId = v.ventureId and v.standalone = 0
	join sGuideNomenk n on n.nomNom = old_name.nomNom
	where numOrder = old_name.numDoc;


	if remoteServerNew is not null then
		set v_cenaEd = new_name.intQuant;
		set v_quantity = round(new_name.quantity/v_perList, 2);
		if update(quantity) or update(intQuant) then
			call update_remote(remoteServerNew
				, 'scet'
				, 'summa_sale'
				, convert(varchar(20), v_quantity * v_cenaEd)
				, 'id = ' + convert(varchar(20), v_id_scet)
			);
        end if;
		
		if update(quantity) then
			call update_remote(
				remoteServerNew
				, 'scet'
				, 'kol1'
				, convert(varchar(20), v_quantity)
				, 'id = ' + convert(varchar(20), v_id_scet)
			);
		end if;
	end if;
	  
end;
	
	
if exists (select 1 from systriggers where trigname = 'wf_delete_nomenk' and tname = 'sDmcRez') then 
	drop trigger sDmcRez.wf_delete_nomenk;
end if;
    
create TRIGGER "wf_delete_nomenk" before delete on
sDmcRez
referencing old as old_name
for each row
begin
	declare remoteServerNew varchar(32);
	
	select 
		sysname
	into 
		remoteServerNew
	from BayOrders o
	join GuideVenture v on o.ventureId = v.ventureId and v.standalone = 0
	where numOrder = old_name.numDoc;

	if remoteServerNew is not null then
		call delete_remote(remoteServerNew, 'scet', 'id = ' + convert(varchar(20), old_name.id_scet));
	end if;
end;



----------------------------------------------------------------------
--------------         BayNomenkOut          -----------------
----------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_BayNomenkOut_outcome_di' and tname = 'BayNomenkOut') then 
	drop trigger BayNomenkOut.wf_BayNomenkOut_outcome_di;
end if;
/*
create 
	trigger wf_BayNomenkOut_outcome_di before delete order 1 on 
BayNomenkOut
referencing old as old_name
for each row
begin
--	declare v_id_mat integer;
	declare v_id_jmat integer;
	declare v_sysname varchar(50);

--	set v_id_mat = old_name.id_mat;
	set v_id_jmat = old_name.id_jmat;

	select v.sysname
	into v_sysname
	from BayOrders o
		left join guideventure v on v.ventureid = o.ventureid and v.standalone = 0
	where numorder = old_name.numOrder;

	call wf_otgruz_remove (
		v_id_jmat
		,'stime'
	);

	if v_sysname is not null and v_sysname != 'stime' then
		call wf_otgruz_remove (
			v_id_jmat
			,v_sysname
		);

	end if;

		
		
end;

*/



if exists (select 1 from systriggers where trigname = 'wf_BayNomenkOut_outcome_ui' and tname = 'BayNomenkOut') then 
	drop trigger BayNomenkOut.wf_BayNomenkOut_outcome_ui;
end if;

/*

create 
	trigger wf_BayNomenkOut_outcome_ui before update order 1 on 
BayNomenkOut
referencing new as new_name old as old_name
for each row
begin
	declare v_id_mat integer;
	declare v_id_jmat integer;
	declare v_sysname varchar(50);
	declare v_cena float;

	if update(quant) and old_name.quant != new_name.quant then
		set v_id_mat = old_name.id_mat;
		set v_id_jmat = old_name.id_jmat;

		select intQuant into v_cena from sDmcRez where numOrder = new_name.numOrder and nomnom = new_name.nomNom;

		select v.sysname
		into v_sysname
		from BayOrders o
		left join guideventure v on v.ventureid = o.ventureid and v.standalone = 0
		where numorder = old_name.numOrder;
		

		call wf_otgruz_quant(
			v_id_mat
			,v_id_jmat
			,new_name.quant
			,v_cena
			,'stime'
		);

		if v_sysname is not null and v_sysname != 'stime' then
			call wf_otgruz_quant(
				v_id_mat
				,v_id_jmat
				,new_name.quant
				,v_cena
				,v_sysname
			);

		end if;


	end if;
end;
*/


if exists (select 1 from systriggers where trigname = 'wf_BayNomenkOut_outcome_bi' and tname = 'BayNomenkOut') then 
	drop trigger BayNomenkOut.wf_BayNomenkOut_outcome_bi;
end if;
/*
create 
	trigger wf_BayNomenkOut_outcome_bi before insert order 1 on 
BayNomenkOut
referencing new as new_name
for each row
begin
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_jmat_nu varchar(20);
	declare v_mat_nu varchar(20);
	declare v_currency_rate float;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_id_source integer;
	declare v_id_dest integer;
--	declare v_osn varchar(100);
	declare v_id_jscet integer;
--	declare v_venture_id integer;
	declare v_firm_id integer;
	declare v_sysname varchar(50);
	declare v_ventureName varchar(100);
	declare v_cena float;
	declare v_cur_otgruz_date date;



	select max(id_jmat) into v_id_jmat 
	from BayNomenkOut 
	where numOrder = new_name.numOrder and outDate = new_name.outDate;

	select 
		 o.id_jscet
		, isnull(s.id_voc_names, 0)
		, isnull(f.id_voc_names,0)
		, v.ventureName
		, v.sysname
	into  
		 v_id_jscet
		, v_id_source
		, v_id_dest
		, v_ventureName
		, v_sysname
	from BayOrders o
		left join guideventure v on v.ventureid = o.ventureid and v.standalone = 0
		left join BayGuideFirms f on o.firmid = f.firmid
		left join sguidesource s on sourceid = -1001
	where numorder = new_name.numOrder;

	
	set v_id_currency = system_currency();
	call slave_currency_rate_stime(v_datev, v_currency_rate);

--	select id_voc_names into v_id_dest from guidefirms where firmid = v_firm_id;
--	    message 'v_id_dest = ', v_id_dest to client;
	-- со склада 1 
	-- ?? хотя по идее нужно бы отгружать со склада готовой продукции
--	select id_voc_names into v_id_source from sguidesource where sourceid = -1001;

	if v_id_jmat is null then
--	    message '---' to client;
		set v_id_jmat = wf_otgruz_jmat(
			new_name.numOrder
			, v_id_jscet
--			, v_venture_id
			, new_name.outDate
			, v_id_source
			, v_id_dest
			, v_id_currency
			, v_datev
			, v_currency_rate
			, v_sysname
		);
--		update BayOrders set id_jmat = v_id_jmat where numorder = new_name.numOrder;
	end if;

	set v_id_mat = get_nextid('mat');
	call slave_select_stime(v_mat_nu, 'mat', 'max(nu)', 'id_jmat = ' + convert(varchar(20), v_id_jmat));
	set v_mat_nu = convert(varchar(20), convert(integer, isnull(v_mat_nu, 0)) + 1);
	select intQuant into v_cena from sDmcRez where numDoc = new_name.numOrder and nomnom = new_name.nomNom;

	call wf_otgruz_nom(
		  v_id_mat
		, v_id_jmat
		, new_name.nomnom
		, new_name.quant
		, v_cena
		, v_mat_nu
		, v_id_source
		, v_id_dest
		, v_currency_rate
		, v_sysname
	);
	set new_name.id_mat = v_id_mat;
	set new_name.id_jmat = v_id_jmat;

end;

*/



----------------------------------------------------------------------
--------------         xUslugOut          -----------------
----------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_xUslugOut_outcome_di' and tname = 'xUslugOut') then 
	drop trigger xUslugOut.wf_xUslugOut_outcome_di;
end if;

/*
create 
	trigger wf_xUslugOut_outcome_di before delete order 1 on 
xUslugOut
referencing old as old_name
for each row
begin
--	declare v_id_mat integer;
	declare v_id_jmat integer;
	declare v_sysname varchar(50);

--	set v_id_mat = old_name.id_mat;
	set v_id_jmat = old_name.id_jmat;

	select v.sysname
	into v_sysname
	from Orders o
		left join guideventure v on v.ventureid = o.ventureid and v.standalone = 0
	where numorder = old_name.numOrder;

	call wf_otgruz_remove (
		v_id_jmat
		,'stime'
	);

	if v_sysname is not null and v_sysname != 'stime' then
		call wf_otgruz_remove (
			v_id_jmat
			,v_sysname
		);

	end if;

		
		
end;

*/



if exists (select 1 from systriggers where trigname = 'wf_xUslugOut_outcome_ui' and tname = 'xUslugOut') then 
	drop trigger xUslugOut.wf_xUslugOut_outcome_ui;
end if;

/*
create 
	trigger wf_xUslugOut_outcome_ui before update order 1 on 
xUslugOut
referencing new as new_name old as old_name
for each row
begin
	declare v_id_mat integer;
	declare v_id_jmat integer;
	declare v_sysname varchar(50);
	declare v_cena float;

	if update(quant) and old_name.quant != new_name.quant then
		set v_id_mat = old_name.id_mat;
		set v_id_jmat = old_name.id_jmat;

--		select intQuant into v_cena from sDmcRez where numOrder = new_name.numOrder and nomnom = new_name.nomNom;

		select v.sysname
			,o.ordered
		into v_sysname
			, v_cena
		from Orders o
		left join guideventure v on v.ventureid = o.ventureid and v.standalone = 0
		where numorder = old_name.numOrder;
		

		call wf_otgruz_quant(
			v_id_mat
			,v_id_jmat
			,new_name.quant
			,v_cena
			,'stime'
		);

		if v_sysname is not null and v_sysname != 'stime' then
			call wf_otgruz_quant(
				v_id_mat
				,v_id_jmat
				,new_name.quant
				,v_cena
				,v_sysname
			);

		end if;


	end if;
end;
*/




if exists (select 1 from systriggers where trigname = 'wf_xUslugOut_outcome_bi' and tname = 'xUslugOut') then 
	drop trigger xUslugOut.wf_xUslugOut_outcome_bi;
end if;

/*
create 
	trigger wf_xUslugOut_outcome_bi before insert order 1 on 
xUslugOut
referencing new as new_name
for each row
begin
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_jmat_nu varchar(20);
	declare v_mat_nu varchar(20);
	declare v_currency_rate float;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_id_source integer;
	declare v_id_dest integer;
--	declare v_osn varchar(100);
	declare v_id_jscet integer;
--	declare v_venture_id integer;
	declare v_firm_id integer;
	declare v_sysname varchar(50);
	declare v_ventureName varchar(100);
	declare v_cena float;
	declare v_cur_otgruz_date date;
	declare v_nomnom char(10);



	set v_nomnom = 'УСЛ';

	select max(id_jmat) into v_id_jmat 
	from xUslugOut 
	where numOrder = new_name.numOrder and outDate = new_name.outDate;

	select 
		 o.id_jscet
		, isnull(s.id_voc_names, 0)
		, isnull(f.id_voc_names,0)
		, v.ventureName
		, v.sysname
		, o.ordered
	into  
		 v_id_jscet
		, v_id_source
		, v_id_dest
		, v_ventureName
		, v_sysname
		, v_cena
	from Orders o
		left join guideventure v on v.ventureid = o.ventureid and v.standalone = 0
		left join BayGuideFirms f on o.firmid = f.firmid
		left join sguidesource s on sourceid = -1001
	where numorder = new_name.numOrder;

	
	set v_id_currency = system_currency();
	call slave_currency_rate_stime(v_datev, v_currency_rate);

--	select id_voc_names into v_id_dest from guidefirms where firmid = v_firm_id;
--	    message 'v_id_dest = ', v_id_dest to client;
	-- со склада 1 
	-- ?? хотя по идее нужно бы отгружать со склада готовой продукции
--	select id_voc_names into v_id_source from sguidesource where sourceid = -1001;

	if v_id_jmat is null then
--	    message '---' to client;
		set v_id_jmat = wf_otgruz_jmat(
			new_name.numOrder
			, v_id_jscet
--			, v_venture_id
			, new_name.outDate
			, v_id_source
			, v_id_dest
			, v_id_currency
			, v_datev
			, v_currency_rate
			, v_sysname
		);
	end if;

	set v_id_mat = get_nextid('mat');
	call slave_select_stime(v_mat_nu, 'mat', 'max(nu)', 'id_jmat = ' + convert(varchar(20), v_id_jmat));
	set v_mat_nu = convert(varchar(20), convert(integer, isnull(v_mat_nu, 0)) + 1);
--	select intQuant into v_cena from sDmcRez where numDoc = new_name.numOrder and nomnom = new_name.nomNom;

	call wf_otgruz_nom(
		  v_id_mat
		, v_id_jmat
		, v_nomnom
		, new_name.quant
		, v_cena
		, v_mat_nu
		, v_id_source
		, v_id_dest
		, v_currency_rate
		, v_sysname
	);
	set new_name.id_mat = v_id_mat;
	set new_name.id_jmat = v_id_jmat;

end;

*/



//===============================================
//    Процедуры обеспечения живучести программ
//===============================================

if exists (select 1 from sysprocedure where proc_name = 'get_standalone') then
	drop function get_standalone;
end if;



CREATE function get_standalone(
	 p_server varchar(50)
	 ,p_remote integer default 0
) returns integer
begin
	declare v_check varchar(23);

	if isnumeric(p_server)=1 then
		select standalone into v_check from guideVenture where ventureId = p_server;
	else
		select standalone into v_check from guideVenture where sysname = p_server;
	end if;
	if v_check is null then
		set get_standalone = 1;
	else 
		set get_standalone = v_check;
	end if;
end;



if exists (select 1 from sysprocedure where proc_name = 'slave_set_standalone') then
	drop function slave_set_standalone;
end if;

// return 1 - successful changing
//		  0 - failed

CREATE function slave_set_standalone(
	 p_status varchar(23)
	 ,p_server varchar(50) default null
	 ,p_remote integer default 0
) returns integer
begin
	set slave_set_standalone = 1;
	if isnumeric(p_server)=1 then
		update guideVenture set standalone = p_status where ventureId = p_server;
	else
		update guideVenture set standalone = p_status where sysname = p_server;
	end if;
	if p_remote = 1 and p_server is not null then
		execute immediate 'call slave_set_standalone_'+ p_server +'( slave_set_standalone, ''' + p_status + ''')';
//		call call_remote(p_server, 'set_standalone', ''''+ p_status + '''');
	end if; 
	exception when others then
		set slave_set_standalone = 0;
end;



if exists (select 1 from sysprocedure where proc_name = 'get_standalone_remote') then
	drop function get_standalone_remote;
end if;

CREATE function get_standalone_remote(
	 p_server varchar(50) default null
) returns integer
begin
	set get_standalone_remote = 0;
	execute immediate 'call slave_get_standalone_'+ p_server +'( get_standalone_remote)';
	exception when others then
		set get_standalone_remote = -1;
end;


-------------------------------------------------------------------------
--------------             BayGuideFirms      ----------------------
-------------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_delete_firm' and tname = 'BayGuideFirms') then 
	drop trigger BayGuideFirms.wf_delete_firm;
end if;

create TRIGGER wf_delete_firm before delete on
BayGuideFirms
referencing old as old_name
for each row
begin
	if old_name.id_voc_names is not null then
		call delete_host('voc_names', 'id = ' + convert(varchar(20), old_name.id_voc_names));
	end if;
end;


if exists (select 1 from systriggers where trigname = 'wf_update_firm' and tname = 'BayGuideFirms') then 
	drop trigger BayGuideFirms.wf_update_firm;
end if;

create TRIGGER wf_update_firm before update on
BayGuideFirms
referencing old as old_name new as new_name
for each row
begin
	if update(phone) then
		call update_host('voc_names', 'phone', '''''' + new_name.phone + '''''', 'id = ' + convert(varchar(20), old_name.id_voc_names));
	end if;
	if update(fio) then 
		call update_host('voc_names', 'rem', '''''' + new_name.fio + '''''', 'id = ' + convert(varchar(20), old_name.id_voc_names));
	end if;
end;


if exists (select 1 from systriggers where trigname = 'wf_insert_firm' and tname = 'BayGuideFirms') then 
	drop trigger BayGuideFirms.wf_insert_firm;
end if;

create TRIGGER wf_insert_firm before insert on
BayGuideFirms
referencing new as new_name
for each row
begin
	declare v_zakaz_id integer;
	declare v_params varchar(2000);
	declare v_firms_id integer;

	select id_voc_names into v_zakaz_id from BayGuideFirms where firmid = 0;

	-- id  фирмы в базе Комтеха
	set v_firms_id = get_nextid ('voc_names');
	set v_params =
		 convert(varchar(20), v_firms_id)
		+ ', '''''+ substring(new_name.name,1,203) + ''''''
	;
	set v_params = v_params + ', ' + convert(varchar(20), v_zakaz_id);

	call insert_host('voc_names', 'id, nm, belong_id', v_params);

	set new_name.id_voc_names = v_firms_id;
	
end;

-------------------------------------------------------------------------
--------------             GuideFirms      ----------------------
-------------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_delete_firm' and tname = 'GuideFirms') then 
	drop trigger GuideFirms.wf_delete_firm;
end if;

create TRIGGER wf_delete_firm before delete on
GuideFirms
referencing old as old_name
for each row
begin
	if old_name.id_voc_names is not null then
		call delete_host('voc_names', 'id = ' + convert(varchar(20), old_name.id_voc_names));
	end if;
end;


if exists (select 1 from systriggers where trigname = 'wf_update_firm' and tname = 'GuideFirms') then 
	drop trigger GuideFirms.wf_update_firm;
end if;

create TRIGGER wf_update_firm before update on
GuideFirms
referencing old as old_name new as new_name
for each row
begin
	if update(phone) then
		call update_host('voc_names', 'phone', '''''' + new_name.phone + '''''', 'id = ' + convert(varchar(20), old_name.id_voc_names));
	end if;
	if update(fio) then 
		call update_host('voc_names', 'rem', '''''' + new_name.fio + '''''', 'id = ' + convert(varchar(20), old_name.id_voc_names));
	end if;
end;


if exists (select 1 from systriggers where trigname = 'wf_insert_firm' and tname = 'GuideFirms') then 
	drop trigger GuideFirms.wf_insert_firm;
end if;

create TRIGGER wf_insert_firm before insert on
GuideFirms
referencing new as new_name
for each row
begin
	declare v_zakaz_id integer;
	declare v_params varchar(2000);
	declare v_firms_id integer;

	select id_voc_names into v_zakaz_id from guidefirms where firmid = 0;

	-- id  фирмы в базе Комтеха
	set v_firms_id = get_nextid ('voc_names');
	set v_params =
		 convert(varchar(20), v_firms_id)
		+ ', '''''+ substring(new_name.name,1,203) + ''''''
	;
	set v_params = v_params + ', ' + convert(varchar(20), v_zakaz_id);

	call insert_host('voc_names', 'id, nm, belong_id', v_params);

	set new_name.id_voc_names = v_firms_id;
	
end;

-------------------------------------------------------------------------
--------------             sGuideSource      ----------------------
-------------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_delete_source' and tname = 'sGuideSource') then 
	drop trigger sGuideSource.wf_delete_source;
end if;

create TRIGGER wf_delete_source before delete on
sGuideSource
referencing old as old_name
for each row
begin
	if old_name.id_voc_names is not null then
		call delete_host('voc_names', 'id = ' + convert(varchar(20), old_name.id_voc_names));
	end if;
end;


if exists (select 1 from systriggers where trigname = 'wf_update_source' and tname = 'sGuideSource') then 
	drop trigger sGuideSource.wf_update_source;
end if;

create TRIGGER wf_update_source before update on
sGuideSource
referencing old as old_name new as new_name
for each row
begin
	if update(phone) then
		call update_host('voc_names', 'phone', '''''' + new_name.phone + '''''', 'id = ' + convert(varchar(20), old_name.id_voc_names));
	end if;
	if update(fio) then 
		call update_host('voc_names', 'rem', '''''' + new_name.fio + '''''', 'id = ' + convert(varchar(20), old_name.id_voc_names));
	end if;
end;


if exists (select 1 from systriggers where trigname = 'wf_insert_source' and tname = 'sGuideSource') then 
	drop trigger sGuideSource.wf_insert_source;
end if;

create TRIGGER wf_insert_source before insert on
sGuideSource
referencing new as new_name
for each row
begin
	declare v_postav_id integer;
	declare v_params varchar(2000);
	declare v_sources_id integer;

	select id_voc_names into v_postav_id from sGuideSource where sourceid = 0;

	-- id  фирмы в базе Комтеха
	set v_sources_id = get_nextid ('voc_names');
	set v_params =
		 convert(varchar(20), v_sources_id)
		+ ', '''''+ substring(new_name.sourceName,1,203) + ''''''
	;
	set v_params = v_params + ', ' + convert(varchar(20), v_postav_id);

	call insert_host('voc_names', 'id, nm, belong_id', v_params);

	set new_name.id_voc_names = v_sources_id;
	
end;




/*
----------------------------------------------------------------------
--------------                 sDocsVenture          ------------------------
----------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_delete' and tname = 'sDocsVenture') then 
	drop trigger sDocsVenture.wf_delete;
end if;

create TRIGGER wf_delete before delete on
sDocsVenture
referencing old as old_name
for each row
begin
	declare remoteServer varchar(32);
	declare no_echo integer;

	set no_echo = 0;

  	begin
		select @stime into no_echo; 
	exception 
		when other then
			set no_echo = 0;
	end;

	if no_echo = 1 then
		return;
	end if;



	if (old_name.id_jmat is not null) then
		call block_remote('stime', @@servername, 'jmat');
		call block_remote('stime', @@servername, 'mat');
		call delete_remote('stime', 'jmat', 'id = ' + convert(varchar(20), old_name.id_jmat));
		call unblock_remote('stime', @@servername, 'jmat');
		call unblock_remote('stime', @@servername, 'mat');
	end if;

end;



if exists (select 1 from systriggers where trigname = 'wf_insert' and tname = 'sDocsVenture') then 
	drop trigger sDocsVenture.wf_insert;
end if;

create 
	trigger wf_insert before insert on 
sDocsVenture
referencing new as new_name
for each row
begin
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_jmat_nu varchar(20);
	declare v_currency_rate float;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_id_source integer;
	declare v_id_dest integer;
	declare v_osn varchar(100);
	declare v_id_guide_jmat integer;




		set v_id_jmat = get_nextid('jmat');
		set v_id_currency = system_currency();
		call slave_currency_rate_stime(v_datev, v_currency_rate);
--		set v_jmat_nu = new_name.numdoc;
		select id_voc_names into v_id_source from sguidesource where sourceid = new_name.sourid;
		select id_voc_names into v_id_dest from sguidesource where sourceid = new_name.destid;
		set v_osn = '[Prior: '+ convert(varchar(20), new_name.numdoc) +']';
	    
		call wf_insert_jmat (
			'stime'
			,v_id_guide_jmat
			,v_id_jmat
			,now() --v_jmat_date
			,v_jmat_nu
			,v_osn
			,v_id_currency
			,v_datev
			,v_currency_rate
			,v_id_source
			,v_id_dest
		);
		set new_name.id_jmat = v_id_jmat;


end;

if exists (select 1 from systriggers where trigname = 'wf_update' and tname = 'sDocsVenture') then 
	drop trigger sDocsVenture.wf_update;
end if;

create 
	trigger wf_update before update on 
sDocsVenture
referencing new as new_name old as old_name
for each row
begin
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_jmat_nu varchar(20);
	declare v_currency_rate float;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_id_source integer;
	declare v_id_dest integer;
	declare v_osn varchar(100);

	if update(sourid) then
		select id_voc_names into v_id_source from sguidesource where sourceid = new_name.sourid;
		call slave_update_stime('jmat', 'id_s', convert(varchar(20), v_id_source), 'id = ' + convert(varchar(20), old_name.id_jmat));
	end if;
	if update(destid) then
		select id_voc_names into v_id_dest from sguidesource where sourceid = new_name.destid;
		call slave_update_stime('jmat', 'id_d', convert(varchar(20), v_id_dest), 'id = ' + convert(varchar(20), old_name.id_jmat));
	end if;
	if update(note) then
		set v_osn = '[Prior: '+ new_name.note +']';
		call slave_update_stime('jmat', 'osn', '''' +v_osn + '''', 'id = ' + convert(varchar(20), old_name.id_jmat));
	end if;
end;
*/
