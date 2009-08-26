if exists (select '*' from sysprocedure where proc_name like 'wf_set_invoice_detail') then  
	drop procedure wf_set_invoice_detail;
end if;


create procedure wf_set_invoice_detail (
	p_servername varchar(20)
	, p_id_jscet integer
	, p_numOrder integer
	, p_date date
	, p_rate double
	, p_ndsrate double
)
begin
-- Процедура синхронизирует предметы заказа Приора
-- с предметами счета в бухгалтерской базе комтеха
-- Это нужно сделать, если в заказ сначала 
-- добавть предметы, а только потом назначить предприятие,
-- через которую этот заказ должен пройти.

	declare v_id_scet integer;
	declare v_id_inv integer;
	declare is_variant integer;
	declare v_id_variant integer;
	declare is_uslug integer;
	declare v_quant double;
	declare v_perList double;

	set is_uslug = 1; // предполагаем изначально, что да


	for c_nomenk as n dynamic scroll cursor for
		select 
			  p.nomNom as r_nomNom
			, p.quant as r_quant
			, p.cenaEd as r_cenaEd
		from xPredmetybynomenk p
		where p.numOrder = p_numOrder
	do
	    set is_uslug = 0; -- есть предметы к заказу, значит не услуга

		select id_inv, perList into v_id_inv, v_perList from sGuideNomenk where nomnom = r_nomNom;
		
		set v_id_scet = 
			wf_insert_scet(
				p_servername
				, p_id_jscet
				, v_id_inv
				, r_quant / v_perList
				, r_cenaEd * v_perList
				, p_rate
				, p_ndsrate
			);
		update xPredmetyByNomenk set id_scet = v_id_scet where current of n;

	end for;


	for c_izd as i dynamic scroll cursor for
		select 
			  prId as r_prId
			, prExt as r_prExt
			, quant as r_quant
			, cenaEd as r_cenaEd
		from xPredmetyByIzdelia p
		where p.numOrder = p_numOrder
	do

	    set is_uslug = 0; -- есть предметы к заказу, значит не услуга
		select id_inv into v_id_inv from sGuideProducts where prId = r_prId;

		-- смотрим, является ли изделие вариантным?
		
		select count(*) into is_variant from sVariantPower where productId = r_prId;
		if is_variant = 1 then
			-- ищем и/или добавляем вариант в Inv
			set v_id_variant = wf_get_variant_id(p_numOrder, r_prId, r_prExt);
			select id_inv into v_id_inv 
			from sGuideComplect 
			where 
				id_variant = v_id_variant;
		end if;

		set v_id_scet = 
			wf_insert_scet(
				p_servername
				, p_id_jscet
				, v_id_inv
				, r_quant
				, r_cenaEd
				, p_rate
				, p_ndsrate
			);

		update xPredmetyByIzdelia set id_scet = v_id_scet, id_inv = v_id_inv where current of i;
	end for;  -- цикла по изделиям

	select ordered into v_quant from orders where numorder = p_numOrder;
	if is_uslug = 1 and abs(v_quant) > 0.001 then
		-- ищем товар под названием "услуга"
		select id_inv into v_id_inv from sGuideNomenk where nomNom = 'УСЛ';


		set v_id_scet = 
			wf_insert_scet(
				p_servername
				, p_id_jscet
				, v_id_inv
				, 1 // quant
				, v_quant//r_cenaEd
				, p_rate
				, p_ndsrate
			);

	end if;


end;
