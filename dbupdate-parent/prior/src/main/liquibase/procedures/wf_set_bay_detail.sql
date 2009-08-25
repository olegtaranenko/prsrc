if exists (select '*' from sysprocedure where proc_name like 'wf_set_bay_detail') then  
	drop procedure wf_set_bay_detail;
end if;

create procedure wf_set_bay_detail (
			p_servername varchar(20)
			, p_id_jscet integer
			, p_numOrder integer
			, p_date date
			, in p_rate float
)
begin
-- Процедура синхронизирует предметы bay-заказа Приора
-- с предметами счета в бухгалтерской базе комтеха
-- Это нужно сделать, если в заказ сначала 
-- добавть предметы, а только потом назначить предприятие,
-- через которую этот заказ должен пройти.

	declare v_id_scet integer;
	declare v_id_inv integer;
	declare is_variant integer;
	declare v_id_variant integer;
	declare v_quant float;

	for c_nomenk as nn dynamic scroll cursor for
		select 
			  p.nomNom as r_nomNom
			, p.quantity as r_quantity
			, intQuant as r_cenaEd
		from sDmcRez p
		where p.numDoc = p_numOrder
	do

		select 
			n.id_inv
			, r_quantity / n.perList
		into 
			v_id_inv
			, v_quant
		from 
			sGuideNomenk n
		where
			n.nomNom = r_nomNom;


		set v_id_scet = 
			wf_insert_scet(
				p_servername
				, p_id_jscet
				, v_id_inv
				, v_quant
				, r_cenaEd
				, p_date
				, p_rate
			);
		update sDmcRez set id_scet = v_id_scet where current of nn;

	end for;

end;
