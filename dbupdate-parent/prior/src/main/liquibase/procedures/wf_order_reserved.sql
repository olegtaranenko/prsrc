if exists (select 1 from sysprocedure where proc_name = 'wf_order_reserved') then
	drop procedure wf_order_reserved;
end if;

CREATE procedure wf_order_reserved(
	p_nomnom varchar(20)
	, p_days_start integer default null
	, p_days_end integer  default null
)
begin

	create table #order_list (numorder integer);

	insert into #order_list
	select distinct r.numorder 
	from 
		isumBranRsrv r
	where r.nomnom = p_nomnom
		and date1 between isnull(now() - p_days_start, date1) and isnull(now() - p_days_end, date1)
	;

	select r.numorder, r.nomnom, r.quant, r.date1, r.manager, r.client, r.note, r.werk, r.sm_zakazano as sm_zakazano
	, r.sm_paid, r.scope, r.status
	from orderBranRsrv r
	join #order_list o on o.numorder = r.numorder
--	left join orderSellOrde s on s.numorder = r.numorder
	where r.nomnom = p_nomnom
	order by r.date1 desc;

	drop table #order_list;
end;


