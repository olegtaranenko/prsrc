if exists (select 1 from systriggers where trigname = 'wf_insert_nomenk' and tname = 'xPredmetyByNomenk') then 
	drop trigger xPredmetyByNomenk.wf_insert_nomenk;
end if;

