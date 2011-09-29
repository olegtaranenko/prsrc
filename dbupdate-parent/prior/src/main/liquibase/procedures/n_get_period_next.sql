ALTER FUNCTION "DBA"."n_get_period_next" (
	  p_cur date
	, p_period_type varchar(20) default 'month'
) returns date
begin
	execute immediate 'select dateadd(' + p_period_type + ', 1, p_cur) into n_get_period_next';
end