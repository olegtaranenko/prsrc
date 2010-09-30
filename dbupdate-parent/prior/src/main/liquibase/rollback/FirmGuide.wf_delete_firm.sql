if exists (select 1 from systriggers where trigname = 'wf_delete_firm' and tname = 'FirmGuide') then 
	drop trigger FirmGuide.wf_delete_firm;
end if;

create TRIGGER wf_delete_firm before delete on
FirmGuide
referencing old as old_name
for each row
begin
	if old_name.id_voc_names is not null then
		call update_host('voc_names', 'deleted', '1', 'id = ' + convert(varchar(20), old_name.id_voc_names));
	end if;
end;

