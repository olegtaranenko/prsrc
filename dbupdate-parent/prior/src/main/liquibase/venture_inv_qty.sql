if exists (select 1 from sysprocedure where proc_name = 'venture_inv_qty') then
	drop function venture_inv_qty;
end if;

create 
	-- возвращает остаток по позиции для заданного предприятия.
	function venture_inv_qty (
		  p_nomnom varchar(20)
		, p_venture_id integer
		, p_inventory_date date default null
	) returns float
begin

    if p_nomnom is null or p_venture_id is null then
    	raiserror 17000 'Invalid parameter value..';
    end if;

    if p_inventory_date is null then
    	set p_inventory_date = convert(date, now());
    end if;

	create table #saldo(id integer, debit float, kredit float);

	insert into #saldo (id, debit, kredit)
    select r_ventureid, sum(r_qty * r_kredit) as debit, 0
    from dummy
    	join (	
    		select
    				 quant/k.perlist as r_qty
    				, if (n.sourid <= -1001 and n.destid <= -1001) then 
    						0 
    					else 
    						if n.destid <= -1001 then 
    							1
    						else
    							-1
							endif
    					endif as 
					r_kredit
    				, if (n.sourid <= -1001 and n.destid <= -1001) then 
							null 
    					else 
    						if n.destid <= -1001 then 
    							isnull(n.ventureid, v.ventureid) 
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
        			from sdocs n
    				join sdmc m on n.numdoc = m.numdoc 
    						and n.numext = m.numext 
    				join sguidenomenk k on k.nomnom = m.nomnom
    			    join sguidesource s on s.sourceId = n.sourId
    				join sguidesource d on d.sourceId = n.destId
    				join system sys on 1 = 1
    				join guideventure v on v.id_analytic = sys.id_analytic_default
    				left join orders o on o.numorder = n.numdoc
    				left join bayorders bo on bo.numorder = n.numdoc
				where
    					m.nomnom = p_nomnom
					and convert(date, n.xDate) <= isnull(p_inventory_date, convert(date, n.xDate))
    	) x on 1=1
	group by r_ventureid;

	
	
		
	insert into #saldo (id, debit, kredit)
    select srcVentureId, 0, sum(m.quant / k.perlist) as kredit
			from sdmcventure m
			join sdocsventure n on m.sdv_id = n.id and n.cumulative_id is not null
			join sguidenomenk k on k.nomnom = m.nomnom
			where m.nomnom = p_nomnom
				and n.nDate <= isnull(p_inventory_date, n.nDate)
			group by srcVentureId;

	insert into #saldo (id, debit, kredit)
    select dstVentureId, sum(m.quant / k.perlist) as kredit, 0
			from sdmcventure m
			join sdocsventure n on m.sdv_id = n.id and n.cumulative_id is not null
			join sguidenomenk k on k.nomnom = m.nomnom
			where m.nomnom = p_nomnom
				and n.nDate <= isnull(p_inventory_date, n.nDate)
			group by dstVentureId;

	select sum(debit - kredit) 
	into venture_inv_qty
	from #saldo
	where id = p_venture_id;

	drop table #saldo;
end;

