if exists (select 1 from systriggers where trigname = 'wf_update' and tname = 'OrdersEquip') then 
	drop trigger OrdersEquip.wf_update;
end if;

create TRIGGER wf_update after update order 1 on
OrdersEquip
referencing old as old_name new as new_name
for each row
begin

	declare v_numorder integer;
	declare v_outdatetime datetime;

	set v_numorder = old_name.numorder;

	select max(outdatetime) into v_outdatetime from OrdersEquip where numorder = v_numorder;

	update Orders set outdatetime = v_outdatetime
	where numorder = v_numorder;

end;
