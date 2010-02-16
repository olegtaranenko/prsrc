if exists (select 1 from sysprocedure where proc_name = 'deleteOrderEquip') then
	drop procedure deleteOrderEquip;
end if;


