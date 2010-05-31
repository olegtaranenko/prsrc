if exists (select '*' from sysprocedure where proc_name like 'wf_report_mat_ost') then  
	drop procedure wf_report_mat_ost;
end if;

create procedure wf_report_mat_ost (
	p_with_ost int default 0
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


	create table #nomenk(nomnom varchar(20), quant double null, perList integer null, primary key(nomnom));
	
	insert into #nomenk(nomnom, perList)
	select distinct
		k.nomnom, k.perList
	from sGuideNomenk       k
	where k.web = 'web';

	if p_with_ost = 1 then
		create table #saldo(nomnom varchar(20), debit float null, kredit float null);
    
		create table #itogo(nomnom varchar(20), debit float null, kredit float null);
    
--select * from #nomenk;
    
    
		insert into #saldo (nomnom, debit)
        select m.nomnom, sum(isnull(m.quant, 0)) 
        from sdocs n
		join sdmc m on n.numdoc = m.numdoc and n.numext = m.numext
		join #nomenk k on k.nomnom = m.nomnom
        where n.destId = -1001
		group by m.nomnom;
    
    
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
    
    
    
		update #nomenk set quant = #nomenk.quant - isumBranRsrv.quant
		from isumBranRsrv 
		where isumBranRsrv.nomnom = #nomenk.nomnom;
--select * from #nomenk;
		
		select 
			 n.nomnom, trim(n.cod + ' ' + n.nomname + ' ' + n.size) as nomenk, n.ed_izmer2 
			, n.cod as cod, n.nomname, n.size as size
			, round(n.nowOstatki / n.perlist - 0.499, 0) as qty_fact
			, round(k.quant / n.perlist, 2) as qty_dost
			, wf_breadcrump_klass(n.klassid) as klassname, n.klassid
			, n.cena_W, n.rabbat, n.margin,  n.kolonok, n.CenaOpt2, n.CenaOpt3, n.CenaOpt4
			, g.kolon1, g.kolon2, g.kolon3, g.kolon4
		from     #nomenk          k
			join sGuideNomenk     n on k.nomnom = n.nomnom
			join #sGuideKlass_ord ko on n.klassId = ko.id
			join sGuideKlass      g on g.klassId = n.klassId
		order by ko.ord, n.nomnom;
	else 
		select
			 n.nomnom, trim(n.cod + ' ' + n.nomname + ' ' + n.size) as nomenk, n.ed_izmer2 
			, n.cod as cod, n.nomname, n.size as size
			, round(n.nowOstatki / n.perlist - 0.499, 0) as qty_fact
			, wf_breadcrump_klass(n.klassid) as klassname, n.klassid
			, n.cena_W, n.rabbat, n.margin,  n.kolonok, n.CenaOpt2, n.CenaOpt3, n.CenaOpt4
			, g.kolon1, g.kolon2, g.kolon3, g.kolon4
		from     #nomenk          k
			join sGuideNomenk     n on k.nomnom = n.nomnom
			join #sGuideKlass_ord ko on n.klassId = ko.id
			join sGuideKlass      g on g.klassId = n.klassId
		order by ko.ord, n.nomnom;
	end if;


	drop table #sGuideKlass_ord;



end;
