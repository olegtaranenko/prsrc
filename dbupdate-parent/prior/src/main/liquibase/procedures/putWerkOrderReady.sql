if exists (select 1 from sysprocedure where proc_name = 'putWerkOrderReady') then
	drop procedure putWerkOrderReady;
end if;


CREATE procedure putWerkOrderReady (
	p_numorder integer
	, p_xDate varchar(10)
	, p_obrazec double
	) 
begin
	
	update OrderInCeh set worktime = p_worktime
		, outDatetime = p_outDatetime
	where numorder = p_numorder
		and equipId = p_equipId;

	if @@rowcount = 0 then
		insert into OrdersEquip (numorder, equipId, worktime, outDatetime)
		values (p_numorder, p_equipId, p_worktime, p_outDatetime);
	end if;

end;

