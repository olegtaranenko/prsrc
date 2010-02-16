if exists (select 1 from systriggers where trigname = 'wf_insert' and tname = 'OrdersEquip') then 
	drop trigger OrdersEquip.wf_insert;
end if;

