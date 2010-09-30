if exists (select 1 from systriggers where trigname = 'wf_update_firm' and tname = 'FirmGuide') then 
	drop trigger FirmGuide.wf_update_firm;
end if;


create TRIGGER wf_update_firm before update on
FirmGuide
referencing old as old_name new as new_name
for each row
begin
	if update(phone) then
		call update_host('voc_names', 'phone', '''''' + new_name.phone + '''''', 'id = ' + convert(varchar(20), old_name.id_voc_names));
	end if;
	if update(fio) then 
		call update_host('voc_names', 'rem', '''''' + new_name.fio + '''''', 'id = ' + convert(varchar(20), old_name.id_voc_names));
	end if;
end;

