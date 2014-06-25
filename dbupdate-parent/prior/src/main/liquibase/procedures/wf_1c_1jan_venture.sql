if exists (select '*' from sysprocedure where proc_name like 'wf_1c_1jan_venture') then  
	drop procedure wf_1c_1jan_venture;
end if;

create procedure wf_1c_1jan_venture (
)
begin

	
	declare p_id_name varchar(64);       
	declare p_parent_id_name varchar(64);
	declare p_order_by_name varchar(256);
	declare p_1jan2013 date;

	create table #nomenk(
		ventureId int,
		nomnom varchar(20), 
		quant double null, 
		quantDost double null, 
		perList integer null, 
		primary key (ventureId, nomnom)
	);
	
	insert into #nomenk(nomnom, ventureId)
	select 
		k.nomnom, v.ventureId
	from  sGuideNomenk k,
		GuideVenture v;



	set p_1jan2013 = '20130101';

	call wf_calculate_ost_venture (p_1jan2013);

	select 
		 n.nomnom, trim(n.cod + ' ' + n.nomname + ' ' + n.size) as nomenk, n.ed_izmer2 
		, n.cod as cod, n.nomname, n.size as size
		, isnull(round(k1.quant / n.perlist, 2), 0) as qty_pm
		, isnull(round(k2.quant / n.perlist, 2), 0) as qty_mm
		, isnull(round(k3.quant / n.perlist, 2), 0) as qty_an
		, round((isnull(k1.quant, 0) + isnull(k2.quant, 0) + isnull(k3.quant, 0)) / n.perlist, 2) as total
	from     sGuideNomenk   n 
		join #nomenk        k1 on k1.nomnom = n.nomnom and k1.ventureId = 1
		join #nomenk        k2 on k2.nomnom = n.nomnom and k2.ventureId = 2
		join #nomenk        k3 on k3.nomnom = n.nomnom and k3.ventureId = 3
	where round(total, 2) > 0
	order by n.nomnom;

end;
