-- исправить суммы по межскладским накладным, которые
-- были уже заведены

if exists (select 1 from systriggers where trigname = 'wf_income_detail_upd' and tname = 'mat' and event='UPDATE') then 
	drop trigger mat.wf_income_detail_upd;
end if;


update mat set kol1 = kol2, kol3 = kol2
where kol2 != 0 and kol2 != kol1 
and tp1 = 3 and tp2 = 2 and tp3 = 1 and tp4 = 0
;
commit;

