if exists (select 1 from systriggers where trigname = 'wf_update_nomenk' and tname = 'sDmcRez') then 
	drop trigger sDmcRez.wf_update_nomenk;
end if;

create TRIGGER wf_update_nomenk before update on
sDmcRez
referencing old as old_name new as new_name
for each row
begin
	declare v_id_scet       integer;
	declare v_id_jscet      integer;
	declare v_id_inv        integer;
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
		, o.id_jscet
		, n.id_inv
	into remoteServerOld
		, v_perList 
		, v_currency_rate
		, v_ndsrate
		, v_id_jscet
		, v_id_inv
	from BayOrders o
	join GuideVenture v on o.ventureId = v.ventureId and v.standalone = 0
	join sGuideNomenk n on n.nomNom = old_name.nomNom
	where numOrder = old_name.numDoc;


	if remoteServerOld is not null then
		set v_quantity = new_name.quantity/v_perList;
		if update(quantity) or update(intQuant) then
			set v_updated = wf_scet_price_changed(
				remoteServerOld
				, v_quantity
				, new_name.intQuant
				, old_name.id_scet
				, v_currency_rate
				, v_ndsrate
				, v_id_jscet
				, v_id_inv
			);
			if v_updated > 0 then
				set new_name.id_scet = v_updated;
			end if;
        end if;
		if update(quantity) and old_name.id_scet is not null then
			call update_remote(remoteServerOld, 'scet', 'kol1', convert(varchar(20), v_quantity), 'id = ' + convert(varchar(20), old_name.id_scet));
		end if;
	end if;

end;
	

