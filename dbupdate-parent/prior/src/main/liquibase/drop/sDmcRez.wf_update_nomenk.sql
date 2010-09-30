if exists (select 1 from systriggers where trigname = 'wf_update_nomenk' and tname = 'sDmcRez') then 
	drop trigger sDmcRez.wf_update_nomenk;
end if;

