if exists (select 1 from systriggers where trigname = 'wf_update_izd' and tname = 'xPredmetyByIzdelia') then 
	drop trigger xPredmetyByIzdelia.wf_update_izd;
end if;

create TRIGGER wf_update_izd before update on
xPredmetyByIzdelia
referencing old as old_name new as new_name
for each row
begin
	declare v_id_scet integer;
	declare v_numorder integer;
	declare v_belong_id integer;
	declare remoteServerNew varchar(32);
	declare v_values varchar(100);
	declare v_fields varchar(200);
	declare v_currency_rate double;
	declare v_ndsrate       float;
	
	set v_id_scet = old_name.id_scet;
--	set v_numorder = old_name.numOrder;

	select sysname, rate
		, v.nds
	into remoteServerNew, v_currency_rate
		, v_ndsrate
		from orders o
		join GuideVenture v on o.ventureId = v.ventureId and v.standalone = 0
		where numOrder = old_name.numOrder;


	if remoteServerNew is not null then
		if update(quant) or update(cenaEd) then
			call wf_scet_price_changed(remoteServerNew, new_name.quant, new_name.cenaEd, v_id_scet, v_currency_rate, v_ndsrate)
		end if;
		if update(quant) then
			call update_remote(remoteServerNew, 'scet', 'kol1', convert(varchar(20), new_name.quant), 'id = ' + convert(varchar(20), v_id_scet));
		end if;
	end if;
  
end;

