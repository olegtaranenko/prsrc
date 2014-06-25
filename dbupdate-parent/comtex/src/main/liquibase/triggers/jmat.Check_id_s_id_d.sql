if exists (select 1 from systriggers where trigname = 'Check_id_s_id_d' and tname = 'jmat') then 
	drop trigger jmat.Check_id_s_id_d;
end if;
