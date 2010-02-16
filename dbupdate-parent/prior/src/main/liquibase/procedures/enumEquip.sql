if exists (select 1 from sysprocedure where proc_name = 'enumEquip') then
	drop function enumEquip;
end if;


CREATE function enumEquip (
	p_numorder integer
	) returns varchar(16)
begin
	declare v_first bit;
	declare v_second bit;
	declare v_first_eq varchar(4);
	set v_first = 0; set v_second = 0;

	for x as xc dynamic scroll cursor for
		select h.ceh as r_ceh
		from OrdersEquip oe
		join GuideCeh h on h.cehId = oe.cehId
		where oe.numorder = p_numorder
		order by h.cehId
	do
		if v_first = 0 then
			set v_first = 1;
			set enumEquip = r_ceh;
			set v_first_eq = r_ceh;
		elseif v_second = 0 then
			set v_second = 1;
			set enumEquip = substring(v_first_eq, 1, 1) + '/' + substring(r_ceh, 1, 1)
		else
			set enumEquip = enumEquip + '/' + substring(r_ceh, 1, 1)
		
		end if;
	end for;
end;
