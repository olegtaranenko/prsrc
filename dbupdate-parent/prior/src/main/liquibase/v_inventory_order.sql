if exists (select 1 from sysprocedure where proc_name = 'v_inventory_order') then
	drop procedure v_inventory_order;
end if;

create
-- Процедура инвентаризации по предприятию на дату
-- если первый параметр null - по всем придприятиям
-- если второй параметр null - на текущую дату
	procedure v_inventory_order (
		 p_venture_id integer default null
		, p_inventory_date date default null
		, p_total_start integer default 1
	) 
begin

	declare v_id_inventar integer;
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_fields varchar(200);
	declare v_values varchar(2000);
	declare v_nu varchar(20);
	declare v_mat_nu integer;
	declare v_quant float;
	declare v_currency_rate real;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_osn varchar(100);

	if p_inventory_date is null then
		set p_inventory_date = convert(date, now());
	end if;

	create table #saldo(nomnom varchar(20), id integer, debit float, kredit float);

	create table #itogo(nomnom varchar(20), id integer, debit float, kredit float);

	insert into #saldo (nomnom, id, debit, kredit)
	select r_nomnom, r_ventureid, sum(r_qty * r_kredit) as debit, 0
	from dummy
		join (
			select
				 quant as r_qty
				, m.nomnom as r_nomnom
				, if (n.sourid <= -1001 and n.destid <= -1001) then 
						0 
					else 
						if n.destid <= -1001 then 
							1
						else
							-1
   						endif
    			  endif 
	    			as r_kredit
    			, n.ventureid as r_ventureid
        	from sdocs n
    		join sdmc m on n.numdoc = m.numdoc and n.numext = m.numext 
--    		join sguidenomenk k on k.nomnom = m.nomnom
    		join sguidesource s on s.sourceId = n.sourId
    		join sguidesource d on d.sourceId = n.destId
    		join system sys on 1 = 1
    		join guideventure v on v.id_analytic = sys.id_analytic_default
    		left join orders o on o.numorder = n.numdoc
    		left join bayorders bo on bo.numorder = n.numdoc
			where
    			convert(date, n.xDate) <= isnull(p_inventory_date, convert(date, n.xDate))
    	) x on 1=1
	group by r_nomnom, r_ventureid;

	
	
		
	insert into #saldo (nomnom, id, debit, kredit)
    select m.nomnom, srcVentureId, 0, sum(m.quant) as kredit
			from sdmcventure m
			join sdocsventure n on m.sdv_id = n.id and n.cumulative_id is not null
--			join sguidenomenk k on k.nomnom = m.nomnom
			where n.nDate <= isnull(p_inventory_date, n.nDate)
			group by 
				m.nomnom, srcVentureId;

	insert into #saldo (nomnom, id, debit, kredit)
    select m.nomnom, dstVentureId, sum(m.quant) as kredit, 0
			from sdmcventure m
			join sdocsventure n on m.sdv_id = n.id and n.cumulative_id is not null
--			join sguidenomenk k on k.nomnom = m.nomnom
			where n.nDate <= isnull(p_inventory_date, n.nDate)
			group by 
				m.nomnom, dstVentureId;

	insert into #itogo (nomnom, id, debit, kredit)
	select s.nomnom, id, sum(debit), sum(kredit) 
	from #saldo s
	group by 
		s.nomnom, s.id;


	select id_voc_names into v_id_inventar from sguidesource where sourceName = 'Инвентаризация';
	set v_osn = 'Текущая инвентаризация';
		set v_id_jmat = get_nextid('jmat');

        -- глобальный для загловков накладных
		set v_id_mat = get_nextid('mat');
--		set v_currency_rate = system_currency_rate();
		set v_id_currency = system_currency();
		call slave_currency_rate_stime(v_datev, v_currency_rate);

   	for venture_cur as s dynamic scroll cursor for
		select 
			ventureid as r_ventureid
			, sysname as r_server
			, id_sklad as r_id_sklad
		from guideventure v
		where isnull(v.invCode, '' ) != '' and isnull(p_venture_id, v.ventureid) = v.ventureid
	do
		
			set v_nu = select_remote(r_server, 'jmat', 'max(nu)', 'id_guide = 1023');
			set v_nu = convert(varchar(20), convert(integer, isnull(v_nu, 0)) + 1);


			call wf_insert_jmat (
				r_server
				,'1023' --инветаризация
				,v_id_jmat
				,p_inventory_date
				,v_nu
				,v_osn
				,v_id_currency
				,v_datev
				,v_currency_rate
				,v_id_inventar
				,r_id_sklad
			);

        	-- Добавляем предметы к накладной
        	set v_mat_nu = 1;
			for nom_cur as n dynamic scroll cursor for
				select 
					i.nomnom as r_nomnom
					, n.id_inv as r_nomenklature_id
					, debit as r_debit
					, kredit as r_kredit 
					, n.cost as r_cost
					, n.perlist as r_perlist
				from #itogo i
				join sguidenomenk n on n.nomnom = i.nomnom
	            where i.id = r_ventureid
	            order by n.nomname
			do
				set v_quant = r_debit - r_kredit;

				if v_quant >= 0.01 then

--					select cost, perList into v_cost, v_perList from sguidenomenk where nomnom = r_nomnom;

					call wf_insert_mat (
						r_server
						,v_id_mat
						,v_Id_jmat
						,r_nomenklature_id
						,v_mat_nu
						,v_quant
						,r_cost
						,v_currency_rate
						,v_id_inventar
						,r_id_sklad
						,r_perList
					);

					set v_id_mat = v_id_mat + 1;
					set v_mat_nu = v_mat_nu + 1;
				end if;

			end for;
			set v_id_jmat = v_id_jmat + 1;
	end for;

	drop table #saldo;
	drop table #itogo;
end;
