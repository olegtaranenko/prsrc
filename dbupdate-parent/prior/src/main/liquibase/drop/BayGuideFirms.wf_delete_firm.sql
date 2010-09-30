if exists (select 1 from systriggers where trigname = 'wf_delete_firm' and tname = 'BayGuideFirms') then 
	drop trigger BayGuideFirms.wf_delete_firm;
end if;
