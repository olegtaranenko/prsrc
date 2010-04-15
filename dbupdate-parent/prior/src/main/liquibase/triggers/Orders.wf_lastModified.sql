if exists (select 1 from systriggers where trigname = 'last_modified' and tname = 'Orders') then 
	drop trigger Orders.last_modified;
end if;

if exists (select 1 from systriggers where trigname = 'wf_lastModified' and tname = 'Orders') then 
	drop trigger Orders.wf_lastModified;
end if;

create TRIGGER wf_lastModified before update order 2 on
Orders
referencing new as new_name old as old_name
for each row
begin
	declare do_correction int;

	if not update(rowLock) and not update(numorder) and not update(lastModified) and not update(lastManagId) and not update(id_bill) then
		set do_correction = 1;
		if update(dateRS) then
			if isnull(old_name.dateRS, convert(datetime, '20000101')) != isnull(new_name.dateRS, convert(datetime, '20000101')) then
				set do_correction = 1;
			else
				set do_correction = 0;
			end if;
		end if;
		if do_correction = 1 then
			raiserror 17000 'test raiserror!';
			set new_name.lastModified = now();
			begin
				set new_name.lastManagId = @managerId;
			exception when others then
				set new_name.lastManagId = null;
			end;
		end if;
	end if;
end;
