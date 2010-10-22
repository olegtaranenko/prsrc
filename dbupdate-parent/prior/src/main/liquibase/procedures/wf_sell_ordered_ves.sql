if exists (select '*' from sysprocedure where proc_name like 'wf_sell_ordered_ves') then  
	drop procedure wf_sell_ordered_ves;
end if;




create procedure wf_sell_ordered_ves (
	p_numorder integer
)
begin

	create table #productVes(
		prId int,
		prExt int,
		totVes double
	);

	insert into #productVes (prId, prExt, totVes)
	select 
		  i.prid
		, i.prExt
		, sum(i.quant * n.ves / n.perlist)
	from itemSellOrde i
	join sGuideNomenk n on n.nomnom = i.nomnom
	where i.numorder = p_numorder
		and i.prid is not null and i.prext is not null
	group by i.prid, i.prExt;
	

	SELECT gp.prName, gp.prDescript, i.*, ei.eQuant, ei.prevQuant, pv.totVes
	FROM 		sGuideProducts gp  
	INNER JOIN 	xPredmetyByIzdelia 	i  ON gp.prId = i.prId  
	LEFT JOIN 	xEtapByIzdelia 		ei ON i.prExt = ei.prExt AND i.prId = ei.prId AND i.Numorder = ei.Numorder 
	LEFT JOIN 	#productVes 		pv on pv.prId = i.prId AND pv.prExt = i.prExt
	WHERE i.numorder = p_numorder;

end