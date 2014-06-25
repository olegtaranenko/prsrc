if exists (select '*' from sysprocedure where proc_name like 'wf_1c_1jan') then  
	drop procedure wf_1c_1jan;
end if;

create procedure wf_1c_1jan (
	p_sklad_id int default -1001// -1001 (основной) or -1002 (обрезки)
)
begin

	
	declare p_id_name varchar(64);       
	declare p_parent_id_name varchar(64);
	declare p_order_by_name varchar(256);
	declare p_1jan2013 date;

	create table #nomenk(nomnom varchar(20), quant double null, quantDost double null, perList integer null, primary key(nomnom));
	
	insert into #nomenk(nomnom)
	select 
		k.nomnom
	from  sGuideNomenk k;



	set p_1jan2013 = '20130101';

	call wf_calculate_ost_fact_dost (1, p_1jan2013, p_sklad_id);

	select 
		 n.nomnom, trim(n.cod + ' ' + n.nomname + ' ' + n.size) as nomenk, n.ed_izmer2 
		, n.cod as cod, n.nomname, n.size as size
		, round(n.nowOstatki / n.perlist - 0.499, 0) as qty_fact
		, isnull(round(k.quant / n.perlist, 2), 0) as qty_sklad1
		, isnull(round(k.quantDost / n.perlist, 2), 0) as qty_dost
	from     #nomenk        k
		join sGuideNomenk   n on k.nomnom = n.nomnom
	where qty_sklad1 > 0;


end;
