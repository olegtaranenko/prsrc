if exists (select 1 from sysprocedure where proc_name = 'wf_scet_price_changed') then
	drop function wf_scet_price_changed;
end if;


CREATE function wf_scet_price_changed (
-- апдейтим цены(руб и валютную) в бух базе комтеха при изменении суммы в приоре.
	  p_server_new    varchar(32)
	, p_quant         float
	, p_cenaEd        float
	, p_id_scet       integer
	, p_currency_rate float
)
returns integer
begin
	declare v_updated integer;


	set v_updated = update_count_remote(p_server_new, 'scet', 'summa_sale'
		, convert(varchar(20), p_currency_rate * p_quant * p_cenaEd)
		, 'id = ' + convert(varchar(20), p_id_scet)
	);
	set v_updated = update_count_remote(p_server_new, 'scet', 'summa_salev'
		, convert(varchar(20), p_quant * p_cenaEd)
		, 'id = ' + convert(varchar(20), p_id_scet)
	);
	set wf_scet_price_changed = v_updated;

end;


