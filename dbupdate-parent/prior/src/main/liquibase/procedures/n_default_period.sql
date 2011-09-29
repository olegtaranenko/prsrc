ALTER PROCEDURE "DBA"."n_default_period" (
	  inout p_begin       char(20)
	, inout p_end         char(20)
	,       p_filterId    integer    default null
)
begin
	declare v_rpt_min_date date;
	declare v_rpt_max_date date;

	if p_begin is null or char_length(p_begin) = 0 then
		if p_filterId is not null then
			set v_rpt_min_date = convert(date, n_get_booting_param(p_filterId, 'minDate'));
		end if;
		if v_rpt_min_date is not null then
			set p_begin = v_rpt_min_date;
		else
		-- для продаж: если дата начало - ноль - ищем из таблицы самую раннюю дату 
			select min(indate) into p_begin from bayorders;
		end if;
	end if;

	if p_end is null or char_length(p_end) = 0 then
		if p_filterId is not null then
			set v_rpt_max_date = convert(date, n_get_booting_param(p_filterId, 'maxDate'));
		end if;

		if v_rpt_max_date is not null then
			set p_end = v_rpt_max_date;
		else
		-- для продаж: берем текущее значение
			set p_end = now();
		end if;
	end if;
	
	--message 'p_begin = ', p_begin to client;
	--message 'p_end = ', p_end to client;

end