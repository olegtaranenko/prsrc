if exists (select 1 from systriggers where trigname = 'wf_insert_nomenk' and tname = 'sDmcRez') then 
	drop trigger sDmcRez.wf_insert_nomenk;
end if;

