if exists (select '*' from sysprocedure where proc_name like 'wf_naklad_mat_ost') then  
	drop procedure wf_naklad_mat_ost;
end if;

create procedure wf_naklad_mat_ost (
	p_numorder int
	,p_dost int default 0  // считать с доступными остатками или только фактические?
)
begin

	
	declare p_id_name varchar(64);       
	declare p_parent_id_name varchar(64);
	declare p_order_by_name varchar(256);

	create table #nomenk(nomnom varchar(20), quant double null, quantDost double null, perList integer null, primary key(nomnom));
	
	insert into #nomenk(nomnom)
	select 
		k.nomnom
	from  itemBranOrde k
	where numorder = p_numorder;




	call wf_calculate_ost_fact_dost (p_dost);

	select 
		 n.nomnom, trim(n.cod + ' ' + n.nomname + ' ' + n.size) as nomenk, n.ed_izmer2 
		, n.cod as cod, n.nomname, n.size as size
		, round(n.nowOstatki / n.perlist - 0.499, 0) as qty_fact
		, round(k.quant / n.perlist, 2) as qty_sklad1
		, round(k.quantDost / n.perlist, 2) as qty_dost
	from     #nomenk          k
		join sGuideNomenk     n on k.nomnom = n.nomnom;


end;
