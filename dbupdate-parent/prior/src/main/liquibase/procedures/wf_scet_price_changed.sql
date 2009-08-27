if exists (select 1 from sysprocedure where proc_name = 'wf_scet_price_changed') then
	drop function wf_scet_price_changed;
end if;


CREATE function wf_scet_price_changed (
-- апдейтим цены(руб и валютную) в бух базе комтеха при изменении суммы в приоре.
	  p_server_new    varchar(32)
	, p_quant         double
	, p_cenaEd        double
	, p_id_scet       integer
	, p_currency_rate double
	, p_ndsrate       double
	, p_id_jscet      integer
	, p_id_inv        integer
)
returns integer
begin
	declare v_updated integer;
	declare v_rubleEd double;
	declare v_nds  double;


	if p_id_scet is not null then
		if p_cenaEd > 0 then
			set v_nds = p_ndsrate / 100;

			set v_rubleEd = p_currency_rate * p_cenaEd;

			set v_updated = update_count_remote(p_server_new, 'scet', 'summa_sale'
				, convert(varchar(20), v_rubleEd * p_quant)
				, 'id = ' + convert(varchar(20), p_id_scet)
			);
			set v_updated = update_count_remote(p_server_new, 'scet', 'summa_nds'
				, convert(varchar(20), v_rubleEd * p_quant * v_nds / (1 + v_nds))
				, 'id = ' + convert(varchar(20), p_id_scet)
			);
			set v_updated = update_count_remote(p_server_new, 'scet', 'summa_salev'
				, convert(varchar(20), p_quant * p_cenaEd)
				, 'id = ' + convert(varchar(20), p_id_scet)
			);
			set wf_scet_price_changed = -v_updated;
		else
			set v_updated = delete_count_remote(p_server_new, 'scet'
				, 'id = ' + convert(varchar(20), p_id_scet)
			);
			set wf_scet_price_changed = -v_updated * 2;
		end if;

	else
		set wf_scet_price_changed = 
			wf_insert_scet(
				p_server_new
				, p_id_jscet
				, p_id_inv
				, p_quant 
				, p_cenaEd
				, p_currency_rate
				, p_ndsrate
			);
	end if;

end;


