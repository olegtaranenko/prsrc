if exists (select 1 from sysprocedure where proc_name = 'putOrderEquip') then
	drop procedure putOrderEquip;
end if;


