if exists (select 1 from systriggers where trigname = 'wf_delete' and tname = 'FirmTools') then 
	drop trigger FirmTools.wf_delete;
end if;

create TRIGGER wf_delete after delete order 1 on
FirmTools
referencing old as old_name 
for each row
begin

	update FirmGuide set tools = enumTools(old_name.firmId)
	where firmId = old_name.firmId;

end;
