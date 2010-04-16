if exists (select 1 from systriggers where trigname = 'wf_lastModified_insert' and tname = 'OrdersInCeh') then 
	drop trigger OrdersInCeh.wf_lastModified_insert;
end if;

create TRIGGER wf_lastModified_insert after insert order 2 on
OrdersInCeh
referencing new as new_name
for each row
begin

	update orders o set lastModified = now() where o.numorder = new_name.numorder
end;
