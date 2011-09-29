ALTER PROCEDURE "DBA"."n_fill_periods" (
	  p_begin       date
	, p_end         date
	, p_period_type varchar(20) 
	, p_columnId    integer default 0
)
begin
	declare v_st        date;
	declare v_en        date;
	declare v_cur       date;
	declare v_prev      date;
	declare v_period_st date;
	declare v_period_en date;

	call n_default_period(p_begin, p_end);

	set v_cur = n_get_period_st(p_begin, p_period_type);
	set v_period_st = v_cur;

	set n_fill_periods = 0;
	all_periods:
	loop
		message v_cur to client;

		set v_prev = v_cur;
		if v_period_st < p_begin then
			set v_period_st = p_begin;
		else 
			set v_period_st = v_prev;
		end if;

		set v_cur = n_get_period_next(v_cur, p_period_type);
		if v_cur < p_end then
			set v_period_en = v_cur;
		else
			set v_period_en = p_end;
		end if;	
		
    	insert into #periods (st, en, label, year) values (v_period_st, v_period_en, n_get_label(v_period_st, p_period_type), year(v_period_st)); 
		set n_fill_periods = n_fill_periods + 1;

		if v_cur >= p_end then 
			leave all_periods;
		end if;
	end loop;

	if isnull(p_columnId, 0) != 0 then
		delete from #periods where periodId != p_columnId;
	end if;
end