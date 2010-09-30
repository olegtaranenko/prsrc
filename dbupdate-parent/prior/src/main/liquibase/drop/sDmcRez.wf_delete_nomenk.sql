if exists (select 1 from systriggers where trigname = 'wf_delete_nomenk' and tname = 'sDmcRez') then 
	drop trigger sDmcRez.wf_delete_nomenk;
end if;
    
