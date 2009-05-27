/* * * * * * * * * * * * * * * * * * * * * * * * * * *
 * ��������� ��� ��������� ������������� �� �������
 * * * * * * * * * * * * * * * * * * * * * * * * * * */

if exists (select '*' from sysprocedure where proc_name like 'wf_cost_bulk_change') then  
	drop function wf_cost_bulk_change;
end if;


create 
-- �������� ���������� ����������� ���� ��� ������ ������������
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
				-- ��������� � ���� ������ ����������� ������ � sPriceHistory
				select max(change_date) into v_timestamp from sPriceHistory where nomnom = r_nomnom;
				
				update sPriceHistory set bulk_id = v_price_bulk_id where change_date = v_timestamp and nomnom = r_nomnom;
	    
			end if;
		end if;
	end for;

	drop table #tmp_klass;

	return v_price_bulk_id;

end;

