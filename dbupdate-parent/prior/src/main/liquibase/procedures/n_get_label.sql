ALTER FUNCTION "DBA"."n_get_label" (
	  p_st date
	, p_period_type varchar(20) default 'month'
) returns varchar(20)
begin
	if p_period_type = 'month' then
		set n_get_label = substring(convert(varchar(20), p_st, 112), 5, 2);
	elseif p_period_type = 'year' then
		set n_get_label = substring(convert(varchar(20), p_st, 112), 3, 2);
	else
		execute immediate 'select datepart(' + p_period_type + ', p_st) into n_get_label';
	end if;
end