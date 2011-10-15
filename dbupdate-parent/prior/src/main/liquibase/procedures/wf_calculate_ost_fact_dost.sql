if exists (select '*' from sysprocedure where proc_name like 'wf_calculate_ost_fact_dost') then  
	drop procedure wf_calculate_ost_fact_dost;
end if;

create procedure wf_calculate_ost_fact_dost (
	p_dost int default 0
)
/*
	На входе этой процедуры должна быть подготовлена и наполнена таблица 
	create table #nomenk(
		nomnom varchar(20), 
		quant double null, 
		quantDost double null, 
		perList integer null, 
		primary key(nomnom)
	);
	
*/
begin

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
    
    if p_dost = 1 then
		update #nomenk set 
			quantDost = quant;

		update #nomenk set 
			quantDost = #nomenk.quant - isumBranRsrv.quant
		from isumBranRsrv 
		where isumBranRsrv.nomnom = #nomenk.nomnom;
--select * from #nomenk;
	end if;
end;
