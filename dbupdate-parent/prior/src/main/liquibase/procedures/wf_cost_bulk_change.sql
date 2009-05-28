/* * * * * * * * * * * * * * * * * * * * * * * * * * *
 * Процедуры для получения себестоимости из комтеха
 * * * * * * * * * * * * * * * * * * * * * * * * * * */

if exists (select '*' from sysprocedure where proc_name like 'wf_cost_bulk_change') then
	drop function wf_cost_bulk_change;
end if;


create 
-- массовое обновление фактической цены для группы номенклатуры
	function wf_cost_bulk_change (
	p_klassid integer
	, p_cur_rate float default null
) returns integer
begin
	declare v_lvl integer;
	declare v_price_bulk_Id integer;
	declare v_comtex_cost float;
	declare v_timestamp datetime;
	declare v_cur_rate float;
	-- показывает, было ли у позиции движение.
	declare v_has_naklad integer;
	-- кол-во позиций, по которым движения не было.
	declare v_reseted_nomnom integer;

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
	set v_reseted_nomnom = 0;

	for v_table as b1 dynamic scroll cursor for
		select nomnom as r_nomnom, id_inv as r_id_inv
			, cost as r_prior_cost, perList as r_perlist
		from sguidenomenk n
		join #tmp_klass t on n.klassid = t.id
		where id_inv is not null
	do 
		set v_comtex_cost = 0; set v_has_naklad = 0;
		call wf_calc_cost_stime(v_comtex_cost, v_has_naklad, r_id_inv);
		message 'Nomnom =', r_nomnom, ', v_comtex_cost = ', v_comtex_cost, ', v_has_naklad = ',v_has_naklad to client;
		if v_has_naklad = 0 then
			set v_comtex_cost = 0;
		else 
			set v_comtex_cost = v_comtex_cost / v_cur_rate;
		end if;
		if abs(round((v_comtex_cost - r_prior_cost), 2) ) > 0.01 then
    
			-- триггером в этот момент добавляется запись в sPriceHistory
			update sguidenomenk set cost = round(v_comtex_cost, 2) where nomnom = r_nomnom;
			if v_has_naklad > 0 and v_comtex_cost > 0 then
				if v_price_bulk_Id is null then
					insert into sPriceBulkChange (guide_klass_id) values (p_klassid);
					set v_price_bulk_Id = @@identity;
				end if;
				-- обновляем вновь добавленную запись
				select max(change_date) into v_timestamp from sPriceHistory where nomnom = r_nomnom;
				update sPriceHistory set bulk_id = v_price_bulk_id where change_date = v_timestamp and nomnom = r_nomnom;
			else
				set v_reseted_nomnom = v_reseted_nomnom + 1;
			end if;
    
		end if;
		if v_has_naklad = 0 or v_comtex_cost = 0 then
			-- по требованию руководства - удалить историю, еслп не было движения по позиции.
			delete from sPriceHistory where nomnom = r_nomnom;
		end if;
	end for;

	drop table #tmp_klass;

	if isnull(v_price_bulk_id, 0) = 0 then
		return -v_reseted_nomnom;
	else
		return v_price_bulk_id;
	end if;

end;

