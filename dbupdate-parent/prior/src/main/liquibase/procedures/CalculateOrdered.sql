if exists (select 1 from sysprocedure where proc_name = 'CalculateOrdered') then
	drop function CalculateOrdered;
end if;


CREATE function CalculateOrdered (
	p_numorder integer
	) returns float
begin
	select sum(r.quantity / n.perList * intQuant)
	into CalculateOrdered
	from sDmcRez r
	join sGuideNomenk n on r.nomnom = n.nomnom
	where r.numdoc = p_numorder
end;


