if exists (select '*' from sysprocedure where proc_name like 'wf_report_bright_ostat') then  
	drop procedure wf_report_bright_ostat;
end if;

create procedure wf_report_bright_ostat (
	p_prId integer default null 
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


	delete
	from #sGuideSeries_ord 
	where not exists (
		select 1 from sGuideSeries ss, sGuideProducts p where ss.seriaId = #sGuideSeries_ord.id and p.prSeriaId = ss.seriaId and p.prodCategoryId = 2
	);


	create table #nomenk(nomnom varchar(20), quant double null, perList integer null, primary key(nomnom));
	
	create table #saldo(nomnom varchar(20), debit float null, kredit float null);

	create table #itogo(nomnom varchar(20), debit float null, kredit float null);

	create table #products(prId int);

	insert into #products(prId)
	select prId
	from sGuideProducts ph
	where ph.prodCategoryId = 2 and isnumeric(ph.page) = 1 
		and isnull(p_prId, ph.prId) = ph.prId 
	;

	
	delete from #products 
	from sProducts  p 
	where p.productId = #products.prId
		and exists (
			select 1 from sGuideNomenk n
			where n.nomnom = p.nomnom and n.web = 'mat'
		);

--	select * from #products p join sGuideProducts ph on p.prId = ph.prId order by ph.prName;

	insert into #nomenk(nomnom, perList)
	select distinct
		p.nomnom, k.perList
	from #products         tp
	join sGuideProducts    ph on ph.prId = tp.prId
	join #sGuideSeries_ord os on os.id = ph.prSeriaId
	join sProducts          p on p.productId = ph.prId
	join sGuidenomenk       k on k.nomnom = p.nomnom
	where ph.prodCategoryId = 2 and isnumeric(ph.page) = 1 
		and isnull(p_prId, ph.prId) = ph.prId 
	;
--select * from #nomenk;


	insert into #saldo (nomnom, debit)
    select m.nomnom, sum(isnull(m.quant, 0)) 
    from sdocs n
	join sdmc m on n.numdoc = m.numdoc and n.numext = m.numext
	join #nomenk k on k.nomnom = m.nomnom
    where n.destId = -1001
	group by m.nomnom;
--select * from #saldo;


	insert into #saldo (nomnom, kredit)
    select m.nomnom, sum(isnull(m.quant, 0)) 
    from sdocs n
	join sdmc m on n.numdoc = m.numdoc and n.numext = m.numext
	join #nomenk k on k.nomnom = m.nomnom
    where n.sourId = -1001
	group by m.nomnom;

--select * from #saldo;

	insert into #itogo (nomnom, debit, kredit)
    select nomnom, sum(isnull(debit,0)), sum(isnull(kredit,0))
	from #saldo 
    group by nomnom;


--select * from #itogo;

	update #nomenk set quant = #itogo.debit - #itogo.kredit
	from #itogo
	where #itogo.nomnom = #nomenk.nomnom;
--select * from #nomenk;



	update #nomenk set quant = #nomenk.quant - isumBranRsrv.quant
	from isumBranRsrv 
	where isumBranRsrv.nomnom = #nomenk.nomnom;
--select * from #nomenk;


	select 
		  ph.prId, ph.prName, ph.prSeriaId, ph.prSize, ph.prDescript 
		, case isnull(vp.prId, 0) when 0 then 'M' else 'V' end as variative
		, p.xgroup
		, ph.vremObr, ph.formulaNom, ph.cena4, ph.page, ph.sortNom
		, ph.rabbat as productRabbat
    	, wf_breadcrump_seria(ph.prseriaid) as serianame
		, n.nomnom, trim(n.cod + ' ' + n.nomname + ' ' + n.size) as nomenk, ed_izmer2 
		, n.cod as ncod, n.nomname, n.size as nsize
		, round(n.nowOstatki / n.perlist - 0.499, 0) as qty_fact
		, round(k.quant / n.perlist, 2) as qty_dost
		, wf_breadcrump_klass(n.klassid) as klassname, n.klassid
		, n.cena_W, n.rabbat, n.margin,  n.kolonok, n.CenaOpt2, n.CenaOpt3, n.CenaOpt4
		, s.gain2, s.gain3, s.gain4
		, f.formula, w.prId as hasWeb, p.quantity as quantEd
		, s.head1, s.head2, s.head3, s.head4 

	from #products                tp
	join sGuideProducts           ph on tp.prId     = ph.prId
	join #sGuideSeries_ord        os on os.id       = ph.prSeriaId
	left join wf_izdeliaWithWeb    w on w.prId      = ph.prId
	join sProducts                 p on p.productId = ph.prId
	join sGuideNomenk              n on n.nomnom    = p.nomnom
	join sGuideSeries              s on s.seriaId   = ph.prSeriaId
	left join #nomenk              k on k.nomnom    = p.nomnom
	left join sGuideFormuls        f on f.nomer     = ph.formulaNom
	left join vw_VariativeProduct vp on vp.prId   = tp.prId
	where 
			ph.prodCategoryId = 2 
		and isnumeric(ph.page) = 1
		and isnull(p_prId, ph.prId) = ph.prId 
		and isnull(n.web, '') <> 'vmt'
	order by os.ord, ph.sortNom, p.xgroup, n.nomName;


	drop table #sGuideKlass_ord;
	drop table #sGuideSeries_ord;


end;
