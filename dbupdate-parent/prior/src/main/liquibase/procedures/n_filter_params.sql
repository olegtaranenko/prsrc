ALTER PROCEDURE "DBA"."n_filter_params" (
	  p_filterid    integer
)
begin
	select 
		  i.isActive as r_isActive
		, i.id as r_itemId
		, it.itemType as r_itemType
		, it.sqlClause as r_sqlClause
		, p.id as r_paramId
		, intValue as r_intValue
		, charValue as r_charValue
		, pt.paramType as r_paramType
		, pt.paramClass as r_paramClass
		, pt.paramKey as r_paramKey
	from nfilter f
	join nitem i on f.id = i.filterid
	join nItemType it on it.id = i.itemTypeId
	left join nparam p on i.id = p.itemid
	left join nParamType pt on pt.id = p.paramTypeId
	where f.id = p_filterid;
end