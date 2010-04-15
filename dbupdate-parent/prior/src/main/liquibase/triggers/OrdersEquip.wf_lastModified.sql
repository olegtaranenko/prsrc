if exists (select 1 from systriggers where trigname = 'wf_lastModified' and tname = 'OrdersEquip') then 
	drop trigger OrdersEquip.wf_lastModified;
end if;

create TRIGGER wf_lastModified before update order 2 on
OrdersEquip
referencing old as old_name new as new_name
for each row
begin

	if not update(numorder) and not update(lastModified) and not update(lastManagId) then
		set new_name.lastModified = now();
		begin
			set new_name.lastManagId = @managerId;
		exception when others then
			set new_name.lastManagId = null;
		end;
	end if;

end;
