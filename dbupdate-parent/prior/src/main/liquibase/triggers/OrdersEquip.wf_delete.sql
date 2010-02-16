if exists (select 1 from systriggers where trigname = 'wf_delete' and tname = 'OrdersEquip') then 
	drop trigger OrdersEquip.wf_delete;
end if;

create TRIGGER wf_delete after delete order 1 on
OrdersEquip
referencing old as old_name 
for each row
begin
	declare v_orderEquip varchar(16);
	declare v_numorder integer;
	set v_numorder = old_name.numorder;
	set v_orderEquip = enumEquip(v_numorder);
	update Orders set equip = v_orderEquip where numorder = v_numorder;
end;
