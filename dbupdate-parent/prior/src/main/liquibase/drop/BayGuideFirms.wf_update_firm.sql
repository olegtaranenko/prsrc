if exists (select 1 from systriggers where trigname = 'wf_update_firm' and tname = 'BayGuideFirms') then 
	drop trigger BayGuideFirms.wf_update_firm;
end if;
