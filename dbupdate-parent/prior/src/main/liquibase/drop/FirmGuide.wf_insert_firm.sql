if exists (select 1 from systriggers where trigname = 'wf_insert_firm' and tname = 'FirmGuide') then 
	drop trigger FirmGuide.wf_insert_firm;
end if;

