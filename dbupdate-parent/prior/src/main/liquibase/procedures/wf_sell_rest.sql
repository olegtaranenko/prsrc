if exists (select 1 from sysprocedure where proc_name = 'wf_sell_rest') then
	drop procedure wf_sell_rest;
end if;

CREATE procedure wf_sell_rest(
	 p_type     varchar(1)
	,p_numorder varchar(20)
	,p_nomnom   varchar(20)
	,p_prId     integer
	,p_prExt    integer
)
begin
	
	create table #done (quant float);

	if p_type = 'p' then

		select pi.doneQuant, pi.curQuant
		from xPredmetyByIzdelia pi 
		where
			p_prid = pi.prId and p_PrExt = pi.PrExt and pi.numorder = p_numorder;

	elseif p_type = 'n' then

		select sum(isnull(d.quant,0) / n.perlist) as doneQuant, dr.curQuant / n.perlist as curQuant
		from sDmcRez dr  
		join sGuideNomenk n on n.nomnom = dr.nomnom
		left join sdmc d on d.numdoc = dr.numdoc and d.nomnom = dr.nomnom
		where
			dr.nomnom = p_nomnom and dr.numdoc = p_numorder
		group by
			dr.nomnom, dr.curQuant, n.perlist;
	end if;
end;
