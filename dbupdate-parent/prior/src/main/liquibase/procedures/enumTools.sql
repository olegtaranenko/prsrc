if exists (select 1 from sysprocedure where proc_name = 'enumTools') then
	drop function enumTools;
end if;


CREATE function enumTools (
	p_firmId integer
	) returns varchar(32)
begin
	declare v_pass int;
	declare v_tiny_str varchar(32);
	declare v_short_str varchar(32);
	set v_pass = 0;

	for x as xc dynamic scroll cursor for
		select h.ToolShort as r_Tool, h.ToolTiny as r_toolTiny
		from FirmTools oe
		join GuideTool h on h.ToolId = oe.ToolId
		where oe.firmId = p_firmId
		order by h.ToolId
	do
		set v_pass = v_pass + 1;
		if v_pass = 1 then
			set v_short_str = r_Tool;
			set v_tiny_str = r_ToolTiny;
		elseif v_pass = 2 then
			set v_short_str = v_short_str + '+' + r_Tool;
			set v_tiny_str = v_tiny_str + '+' + r_ToolTiny;
		else
			set v_tiny_str = v_tiny_str + '+' + r_ToolTiny;
		end if;
	end for;

	if v_pass <= 2 then
		set enumTools = v_short_str;
	else
		set enumTools = v_tiny_str;
	end if;
end;
