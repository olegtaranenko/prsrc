if exists (select 1 from systriggers where trigname = 'wf_update_nomenk' and tname = 'sDmcRez') then 
	drop trigger sDmcRez.wf_update_nomenk;
end if;

create TRIGGER wf_update_nomenk before update on
sDmcRez
referencing old as old_name new as new_name
for each row
begin
	declare v_id_scet integer;
	declare remoteServerOld varchar(32);

	declare v_cenaEd        double;
	declare v_quantity      double;
	declare v_perList       double;
	declare v_currency_rate double;
	declare v_ndsrate       float;
	declare v_updated       integer;
	
	set v_id_scet = old_name.id_scet;
	  
	select v.sysname
		, n.perList 
		, rate
		, v.nds
	into remoteServerOld
		, v_perList 
		, v_currency_rate
		, v_ndsrate
	from BayOrders o
	join GuideVenture v on o.ventureId = v.ventureId and v.standalone = 0
	join sGuideNomenk n on n.nomNom = old_name.nomNom
	where numOrder = old_name.numDoc;


	if remoteServerOld is not null then
		set v_quantity = new_name.quantity/v_perList;
		if update(quantity) or update(intQuant) then
			set v_updated = wf_scet_price_changed(remoteServerOld, v_quantity, new_name.intQuant, v_id_scet, v_currency_rate, v_ndsrate);
        end if;
		if update(quantity) then
			call update_remote(remoteServerOld, 'scet', 'kol1', convert(varchar(20), v_quantity), 'id = ' + convert(varchar(20), v_id_scet));
		end if;
	end if;

end;
	

