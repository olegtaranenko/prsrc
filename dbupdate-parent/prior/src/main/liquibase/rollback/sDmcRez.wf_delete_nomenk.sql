if exists (select 1 from systriggers where trigname = 'wf_delete_nomenk' and tname = 'sDmcRez') then 
	drop trigger sDmcRez.wf_delete_nomenk;
end if;
    
create TRIGGER wf_delete_nomenk before delete on
sDmcRez
referencing old as old_name
for each row
begin
	declare remoteServerNew varchar(32);
	declare v_id_jscet integer;
	
	select 
		sysname
		, id_jscet
	into 
		remoteServerNew
		, v_id_jscet
	from BayOrders o
	join GuideVenture v on o.ventureId = v.ventureId and v.standalone = 0
	where numOrder = old_name.numDoc;

	if remoteServerNew is not null then
		if old_name.id_scet is not null then
			call delete_remote(remoteServerNew, 'scet', 'id = ' + convert(varchar(20), old_name.id_scet));
			call call_remote(remoteServerNew, 'slave_renu_scet', v_id_jscet);
		end if;
	end if;
end;


