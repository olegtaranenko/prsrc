if exists (select 1 from sysprocedure where proc_name = 'deleteOrderEquip') then
	drop procedure deleteOrderEquip;
end if;


CREATE procedure deleteOrderEquip (
	p_numorder integer
	, p_equipId integer
	) 
begin
	delete from  OrdersEquip 
	where numorder = p_numorder
		and cehId = p_equipId;

end;

