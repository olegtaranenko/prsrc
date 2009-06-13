if exists (select 1 from systriggers where trigname = 'wf_update_nomenk' and tname = 'sDmcRez') then 
	drop trigger sDmcRez.wf_update_nomenk;
end if;

create TRIGGER wf_update_nomenk before update on
sDmcRez
referencing old as old_name new as new_name
for each row
begin
	declare v_id_scet integer;
	declare remoteServerNew varchar(32);

	declare v_cenaEd float;
	declare v_quantity float;
	declare v_perList float;
	declare v_currency_rate float;
	
	set v_id_scet = old_name.id_scet;
	  
	select v.sysname
		, n.perList 
	into remoteServerNew
		, v_perList 
	from BayOrders o
	join GuideVenture v on o.ventureId = v.ventureId and v.standalone = 0
	join sGuideNomenk n on n.nomNom = old_name.nomNom
	where numOrder = old_name.numDoc;


	if remoteServerNew is not null then
		if update(quantity) or update(intQuant) then
			set v_currency_rate = system_currency_rate();
			set v_quantity = round(new_name.quantity/v_perList, 2);
			call update_remote(remoteServerNew, 'scet', 'summa_sale'
				, convert(varchar(20), v_currency_rate * v_quantity * new_name.intQuant)
				, 'id = ' + convert(varchar(20), v_id_scet)
			);
			call update_remote(remoteServerNew, 'scet', 'summa_salev'
				, convert(varchar(20), v_quantity*new_name.intQuant)
				, 'id = ' + convert(varchar(20), v_id_scet)
			);
        end if;
		if update(quantity) then
			call update_remote(remoteServerNew, 'scet', 'kol1', convert(varchar(20), v_quantity), 'id = ' + convert(varchar(20), v_id_scet));
		end if;
	end if;

end;
	

