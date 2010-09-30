if exists (select 1 from systriggers where trigname = 'wf_lastModified' and tname = 'Orders') then 
	drop trigger Orders.wf_lastModified;
end if;
