if exists (select 1 from systriggers where trigname = 'wf_insert' and tname = 'FirmTools') then 
	drop trigger FirmTools.wf_insert;
end if;

create TRIGGER wf_insert after insert order 1 on
FirmTools
referencing new as new_name
for each row
begin
	update FirmGuide set tools = enumTools(new_name.firmId) 
	where  firmId = new_name.firmId;
end;
