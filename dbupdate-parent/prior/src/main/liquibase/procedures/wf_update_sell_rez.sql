if exists (select 1 from sysprocedure where proc_name = 'wf_update_sell_rez') then
	drop procedure wf_update_sell_rez;
end if;

CREATE procedure wf_update_sell_rez(
	 p_numorder varchar(20)
	,p_new_quant float
	,p_old_quant float
	,p_type     varchar(1)
	,p_nomnom   varchar(20)
	,p_prId     integer
	,p_prExt    integer
)
begin
	declare v_vmt     varchar(8);
	declare v_perList float;

	if p_type = 'p' then
		update xPredmetyByIzdelia set curQuant = p_new_quant 
		where numorder = p_numorder and prId = p_prId and prExt = p_prExt;

		update sdmcrez 
			join xPredmetyByIzdelia pi on pi.numorder = p_numorder and pi.prId = p_prId and pi.prExt = p_prExt
			join itemWareOrde ipo on 
				ipo.numorder = p_numorder 
				and ipo.nomnom = sDmcRez.nomnom 
				and ipo.prId = p_prId 
				and ipo.prExt = p_prExt
		set sdmcrez.curQuant = (p_new_quant - p_old_quant) * ipo.quantEd + sdmcrez.curQuant
		where sDmcRez.numdoc = p_numorder;

	elseif p_type = 'n' then
		select web, perList into v_vmt, v_perList from sguidenomenk where nomnom = p_nomnom;
		if v_vmt = 'vmt' then
			set v_perList = 1;
		end if;
		update sDmcRez 
		set curQuant = (p_new_quant - p_old_quant) * v_perList + CurQuant
		where sDmcRez.numdoc = p_numorder and sDmcRez.nomnom = p_nomnom;
	end if;
end;
