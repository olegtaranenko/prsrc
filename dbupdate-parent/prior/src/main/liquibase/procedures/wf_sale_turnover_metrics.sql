if exists (select 1 from sysprocedure where proc_name = 'wf_sale_turnover_metrics') then
	drop function  wf_sale_turnover_metrics;
end if;

create 
	function  wf_sale_turnover_metrics(
	  p_nomnom varchar(20)
	, p_start datetime default null
	, p_end datetime default null
) returns varchar(127)
begin

	declare v_current_total   double;
	declare v_interval_start  date;
	declare v_interval_stop   date;
--	declare v_calculate       integer;
	declare v_saled_quant     double;
	declare v_income_quant    double;
	declare v_current_income  double;
	declare v_outcome_quant   double;
	declare v_period_outcome  integer;
	declare v_base_income_date  date;
	declare v_first_income_date date;
	declare v_prev_sale_date    date;
	declare v_prev_total        double;
	declare o_average_outcome   double;
	declare v_full_period_days  integer;


--	set v_calculate      = 0;
	set v_current_total  = 0;
	set v_income_quant   = 0;
	set v_outcome_quant  = 0;
	set v_period_outcome = 0;
	set v_prev_total     = 0;
	set v_saled_quant    = 0;
--	message 'p_start = ', p_start to client;
--	message 'p_end = ', p_end to client;


	for x as xc dynamic scroll cursor for
		select 
			  i.numdoc as r_numdoc, i.numext as r_numext, i.quant/n.perlist as r_quant
			, d.xdate as r_xdate, d.sourId as r_sourId, d.destId as r_destId
			, b.numorder as r_isBayOrder
		from sdmc i
		join sdocs d on d.numdoc = i.numdoc and d.numext = i.numext
		join sguidenomenk n on n.nomnom = i.nomnom
		left join bayorders b on b.numorder = i.numdoc
		where 
				i.nomnom = p_nomnom
			and d.xdate <= isnull(p_end, d.xdate)
			and ((d.sourId = -1001 or d.destId <= -1001) ) --and not (d.sourId <= -1001 and d.destId <= -1001)
		order by d.xdate, i.numdoc, i.numext
	do
--		message '******************************** ********' to client;
--		message 'r_sourId = ', r_sourId to client;
--		message 'r_destId = ', r_destId to client;
--		message 'r_numdoc = ', r_numdoc to client;
--		message 'r_quant = ', r_quant to client;
--		message 'r_xdate = ', r_xdate to client;


		if p_start is null then
			if v_base_income_date is not null then
				set v_interval_start = v_base_income_date;
			else 
				set v_interval_start = convert(date, r_xdate); -- truncate time.
			end if;
--			message '1) v_interval_start = ', v_interval_start to client;
		else
			if r_xdate >= p_start then
				if p_start < v_base_income_date then
					set v_interval_start = v_base_income_date;
--					message '2) v_interval_start = ', v_interval_start to client;
				else
					set v_interval_start = p_start;
--					message '3) v_interval_start = ', v_interval_start to client;
				end if;
			else
				set v_interval_start = null;
--				message '4) v_interval_start = ', v_interval_start to client;

			end if;
		end if;



		if r_destId = -1001 then
			set v_current_total = v_current_total + r_quant;
			if v_interval_start is not null then
				set v_income_quant = v_income_quant + r_quant;
--				message 'v_income_quant = ', v_income_quant to client;

			end if;
			set v_current_Income = r_quant;

--			message 'v_current_total = ', v_current_total to client;

			if v_prev_total <= 0 then
				set v_base_income_date = r_xdate;
--				message 'v_base_income_date = ', v_base_income_date to client;
			end if;
			if v_first_income_date is null then
				set v_first_income_date = r_xdate;
--				message 'v_first_income_date = ', v_first_income_date to client;
			end if;

		elseif r_sourId <= -1001 then

			set v_current_total = v_current_total - r_quant;
--			message '	v_current_total = ', v_current_total to client;

			if v_interval_start is not null then
				set v_outcome_quant = v_outcome_quant + r_quant;
				if r_isBayOrder is not null then
					set v_saled_quant = v_saled_quant + r_quant;
				end if;
--				set v_income_quant = v_income_quant + v_current_income;
--				message '1) v_outcome_quant = ', v_outcome_quant to client;

			end if;


--			message '   v_interval_start = ', v_interval_start to client;
--			message '	v_interval_stop = ', v_interval_stop to client;
--			message '	v_outcome_quant = ', v_outcome_quant to client;
--			message '	v_period_outcome = ', v_period_outcome to client;
		end if;

		if v_current_total <= 0 then
			if v_interval_start is not null then
				set v_interval_stop = v_prev_sale_date;
--				message '*) v_interval_stop = ', v_interval_stop to client;
--				message '*) v_interval_start = ', v_interval_start to client;
				set v_period_outcome = v_period_outcome + (r_xdate - v_interval_start) + 1;
--				message '1) v_period_outcome = ', v_period_outcome to client;

				set v_interval_start = null;
				set v_interval_stop  = null;
			end if;
		end if;

		set v_prev_sale_date = r_xdate;
		set v_prev_total = v_current_total;
	end for;


	if (v_interval_start is not null and v_interval_stop is null) or (v_current_total > 0) then
		if v_interval_start is null then
--			message '**v_interval_start = ', v_interval_start to client;

			if p_start is null then
				set v_interval_start = v_base_income_date;
			else
				if p_start > v_base_income_date then
					set v_interval_start = p_start;
				else	
					set v_interval_start = v_base_income_date;
				end if;
			end if;
		end if;

		if p_end is null then
			set v_interval_stop = now();
		else
			set v_interval_stop = p_end;
		end if;
--		message '2) v_interval_stop = ', v_interval_stop to client;

		if v_interval_stop >= v_interval_start then
			set v_period_outcome = v_period_outcome + (v_interval_stop - v_interval_start) + 1;
--			message '2) v_period_outcome = ', v_period_outcome to client;
		end if;
	end if;

	if v_period_outcome is not null and v_period_outcome > 0 then
--		message 'v_period_outcome = ', v_period_outcome to client;
--		message 'v_outcome_quant = ', v_outcome_quant to client;
		set o_average_outcome = v_outcome_quant / v_full_period_days * 30;
--		message 'o_average_outcome = ', o_average_outcome to client;

	end if;

	if p_start is null and p_end is null then
		set v_full_period_days = now() - v_first_income_date;
	elseif p_start is null then
		set v_full_period_days = p_end - v_first_income_date;
	elseif p_end is null then
		set v_full_period_days = now() - p_start;
	else 
		set v_full_period_days = p_end - p_start;
	end if;


	set wf_sale_turnover_metrics =
				convert(varchar(20), o_average_outcome) 
		+ ';' + convert(varchar(20), v_full_period_days - v_period_outcome + 1)
		+ ';' + convert(varchar(20), v_saled_quant)
		+ ';' + convert(varchar(20), v_income_quant)
		+ ';' + convert(varchar(20), v_outcome_quant)
	;

end;


