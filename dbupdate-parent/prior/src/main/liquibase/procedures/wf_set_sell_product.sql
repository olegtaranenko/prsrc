if exists (select 1 from sysprocedure where proc_name = 'wf_set_sell_product') then
	drop procedure wf_set_sell_product;
end if;

CREATE procedure wf_set_sell_product(
	 p_numorder integer
	,p_prId     integer
	,p_prExt    integer
	,p_quant    float
)
begin
	update xPredmetyByIzdelia set curQuant = p_quant 
	where numorder = p_numorder and prId = p_prId and prExt = p_prExt;
end;