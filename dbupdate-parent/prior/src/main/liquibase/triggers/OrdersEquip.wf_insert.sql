if exists (select 1 from systriggers where trigname = 'wf_insert' and tname = 'OrdersEquip') then 
	drop trigger OrdersEquip.wf_insert;
end if;

create TRIGGER wf_insert after insert order 1 on
OrdersEquip
referencing new as new_name
for each row
begin
	declare v_orderEquip varchar(16);
	declare v_numorder integer;
	set v_numorder = new_name.numorder;
	set v_orderEquip = enumEquip(v_numorder);
	update Orders set equip = v_orderEquip where numorder = v_numorder;

end;
