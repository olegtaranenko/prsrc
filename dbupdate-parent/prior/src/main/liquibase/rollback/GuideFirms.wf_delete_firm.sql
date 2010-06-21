if exists (select 1 from systriggers where trigname = 'wf_delete_firm' and tname = 'GuideFirms') then 
	drop trigger GuideFirms.wf_delete_firm;
end if;

create TRIGGER wf_delete_firm before delete on
GuideFirms
referencing old as old_name
for each row
begin
	if old_name.id_voc_names is not null then
		call delete_host('voc_names', 'id = ' + convert(varchar(20), old_name.id_voc_names));
	end if;
end;


