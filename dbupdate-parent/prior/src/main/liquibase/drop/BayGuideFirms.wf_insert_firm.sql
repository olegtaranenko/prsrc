if exists (select 1 from systriggers where trigname = 'wf_insert_firm' and tname = 'BayGuideFirms') then 
	drop trigger BayGuideFirms.wf_insert_firm;
end if;
