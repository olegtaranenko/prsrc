if exists (select 1 from sysprocedure where proc_name = 'putOrderEquip') then
	drop procedure putOrderEquip;
end if;


CREATE procedure putOrderEquip (
	p_numorder integer
	, p_equipId integer
	, p_worktime double
	, p_outDatetime datetime
	) 
begin
	update OrdersEquip set worktime = p_worktime
		, outDatetime = p_outDatetime
	where numorder = p_numorder
		and equipId = p_equipId;

	if @@rowcount = 0 then
		insert into OrdersEquip (numorder, equipId, worktime, outDatetime)
		values (p_numorder, p_equipId, p_worktime, p_outDatetime);
	end if;

end;

