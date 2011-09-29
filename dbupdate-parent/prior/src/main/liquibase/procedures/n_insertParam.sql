ALTER FUNCTION "DBA"."n_insertParam" (
	   p_item_id     integer
	, p_param_name   varchar(64)
	, p_intValue     integer default null
	, p_charValue    long varchar default null
) returns integer
begin
	declare v_paramClass varchar(32);
	declare v_varchar long varchar;
	declare v_dateValue date;

	select paramClass 
	into v_paramClass 
	from nParamType pt
	join nItem i           on i.itemTypeId = pt.itemTypeId
	where 
			i.id = p_item_id 
		and pt.paramType = p_param_name;

	if v_paramClass = 'date' then
		set v_dateValue = convert(date, substring(p_charValue, 1, 10), 104);
		set v_varchar = convert(varchar(20), v_dateValue, 104);
	else
		set v_varchar = p_charValue;
	end if;

	insert into nParam (itemId, paramTypeId, intValue, charValue) 
	select p_item_id,  pt.id,  p_intValue, v_varchar
	from 
		  nParamType pt 
		, nItem i
	where 
		pt.paramType = p_param_name 
	and i.id = p_item_id 
	and i.itemTypeId = pt.itemTypeId
	;

	set n_insertParam = @@identity;

end