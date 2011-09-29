ALTER FUNCTION "DBA"."n_get_period_st" (
	  p_begin date
	, p_period_type varchar(20) default 'month'
) returns date
begin
	declare v_cur date;
	declare v_shift_back integer;

	if p_period_type = 'month' then
		set n_get_period_st = ymd(year(p_begin), month(p_begin), 1);
	elseif p_period_type = 'year' then
		set n_get_period_st = ymd(year(p_begin), 1, 1);
	elseif p_period_type = 'week' then

		set v_cur = p_begin;
		set v_shift_back = 0;
		while datepart(dw, v_cur) != 2 loop 		-- понедельник
			set v_cur = dateadd(day, 1, v_cur);
			set v_shift_back = -1;
		end loop;
		set n_get_period_st = dateadd(week, v_shift_back, v_cur);
	else
		set n_get_period_st = p_begin;
	end if;
end