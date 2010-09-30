if exists (select 1 from systriggers where trigname = 'wf_update_orders' and tname = 'Orders') then 
	drop trigger Orders.wf_update_orders;
end if;

