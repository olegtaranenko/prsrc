if exists (select 1 from systriggers where trigname = 'wf_delete' and tname = 'OrdersEquip') then 
	drop trigger OrdersEquip.wf_delete;
end if;

