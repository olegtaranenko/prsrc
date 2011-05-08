if exists (select '*' from sysprocedure where proc_name like 'wf_nomenk_reliz') then  
	drop procedure wf_nomenk_reliz;
end if;

create procedure wf_nomenk_reliz (
	  p_start datetime default null
	, p_end datetime default null
)
begin

	
	declare v_ord_table varchar(64);
	declare p_table_name varchar(64);    
	declare p_id_name varchar(64);       
	declare p_parent_id_name varchar(64);
	declare p_order_by_name varchar(256);
--	create table #klass_ordered (id integer, ord integer);
	set p_table_name = 'sGuideKlass';
	set v_ord_table = get_tmp_ord_table_name(p_table_name);
	execute immediate get_tmp_ord_create_sql(v_ord_table);

	set p_id_name = 'klassId';
	set p_parent_id_name = 'parentKlassId';
	set p_order_by_name = 'klassName';
	call wf_sort_klassificator(p_table_name, p_id_name, p_parent_id_name, p_order_by_name);


	set p_table_name = 'sGuideSeries';
	set v_ord_table = get_tmp_ord_table_name(p_table_name);
	execute immediate get_tmp_ord_create_sql(v_ord_table);

	set p_id_name = 'seriaId';
	set p_parent_id_name = 'parentSeriaId';
	set p_order_by_name = 'seriaName';
	call wf_sort_klassificator(p_table_name, p_id_name, p_parent_id_name, p_order_by_name);

	
	create table #nomenk_reliz (nomnom varchar(20), quant double, sm double);
	create table #izdelia_reliz (prid integer, quant double, sm double, costTotal double);


	insert into #izdelia_reliz
	    select 
    		po.prId, sum(po.quant) as quant, sum(p.cenaEd * po.quant) as cenaTotal, sum(io.costEd * po.quant) as costTotal
		from xpredmetybyizdeliaout po
		join xpredmetybyizdelia p on p.numorder = po.numorder and p.prid = po.prid and p.prext = po.prext
		join orderWareShip io on po.outdate = io.outdate and io.numorder = po.numorder and io.prid = po.prid and io.prext = po.prext
		WHERE po.outDate between isnull(p_start, '20010101') and isnull(p_end, '21001231')
		group by po.prid
	;


	insert into #nomenk_reliz
	    select po.nomnom, sum(po.quant / n.perlist) as quant, sum(p.cenaEd * po.quant) as sum
    	from xpredmetybynomenkout po
		join xpredmetybynomenk p on p.numorder = po.numorder and p.nomnom = po.nomnom
		join sguidenomenk n on n.nomnom = po.nomnom and n.nomnom = p.nomnom
		WHERE po.outDate between isnull(p_start, '20010101') and isnull(p_end, '21001231')
		group by po.nomnom
	;



	select 'Изделия' as outtype, o.ord, trim(g.prDescript + ' ' + g.prsize) as name
			, quant, sm as cenaTotal, convert(varchar(20), r.prid) as id, g.prname as nomnom, g.prSeriaId as klassid
	    	, wf_breadcrump_seria(g.prseriaid) as klassname, 'шт.' as ed_izmer2
    		, r.costTotal / quant as cost -- costEd
		from #izdelia_reliz r
		join sguideproducts g on g.prid = r.prid
		join #sGuideSeries_ord o on o.id = g.prSeriaId
			union
	select 'Номенклатура' as outtype, o.ord, trim(n.cod + ' ' + nomname + ' ' + n.size) as name
    	, s.quant, s.sm as cenaTotal, s.nomnom as id, s.nomnom, k.klassid
		, wf_breadcrump_klass(k.klassid) as klassname, n.ed_izmer2
		, n.cost
	from #nomenk_reliz s
		join sguidenomenk n on n.nomnom = s.nomnom
		join #sGuideKlass_ord o on o.id = n.klassid
    	join sguideklass k on k.klassid = n.klassid
	    order by 1, 2, 3
	;
	
	
	drop table #nomenk_reliz;
	drop table #izdelia_reliz;

	drop table #sGuideKlass_ord;
	drop table #sGuideSeries_ord;

end
	
