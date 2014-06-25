if exists (select '*' from sysprocedure where proc_name like 'wf_calculate_ost_venture') then  
	drop procedure wf_calculate_ost_venture;
end if;

create procedure wf_calculate_ost_venture (
	p_dost int default 0,
	p_inventory_date date default null
)
/*
	На входе этой процедуры должна быть подготовлена и наполнена таблица 
	create table #nomenk(
		ventureId int,
		nomnom varchar(20), 
		quant double null, 
		quantDost double null, 
		perList integer null, 
		primary key(nomnom)
	);
	
*/
begin

		declare v_sklad_id int;
		set v_sklad_id = -1001;

		create table #saldo(nomnom varchar(20), debit float null, kredit float null, ventureId int not null);
    
		create table #itogo(nomnom varchar(20), debit float null, kredit float null, ventureId int not null);
    
--select * from #nomenk;
		
    
		if p_inventory_date is null then
			set p_inventory_date = convert(date, now());
		end if;

		insert into #saldo (nomnom, debit, ventureId)
        select m.nomnom, sum(isnull(m.quant, 0)), n.ventureId
        from sdocs n
		join sdmc m on n.numdoc = m.numdoc and n.numext = m.numext
		join #nomenk k on k.nomnom = m.nomnom
        where n.destId = v_sklad_id
        	and n.xDate < p_inventory_date
		group by m.nomnom, n.ventureId;
    
    
		insert into #saldo (nomnom, kredit, ventureId)
        select m.nomnom, sum(isnull(m.quant, 0)), n.ventureId 
        from sdocs n
		join sdmc m on n.numdoc = m.numdoc and n.numext = m.numext
		join #nomenk k on k.nomnom = m.nomnom
        where n.sourId = v_sklad_id
        	and n.xDate < p_inventory_date
		group by m.nomnom, n.ventureId;
    
--select * from #saldo;
    
		insert into #itogo (nomnom, debit, kredit, ventureId)
        select nomnom, sum(isnull(debit,0)), sum(isnull(kredit,0)), ventureId
		from #saldo 
        group by nomnom, ventureId;
    
    
--select * from #itogo;
		update #nomenk set quant = #itogo.debit - #itogo.kredit
		from #itogo
		where #itogo.nomnom = #nomenk.nomnom and #itogo.ventureId = #nomenk.ventureId;
/*    
    if p_dost = 1 then
		update #nomenk set 
			quantDost = quant;

		update #nomenk set 
			quantDost = #nomenk.quant - isumBranRsrv.quant
		from isumBranRsrv 
		where isumBranRsrv.nomnom = #nomenk.nomnom;
--select * from #nomenk;
	end if;
*/
end;
