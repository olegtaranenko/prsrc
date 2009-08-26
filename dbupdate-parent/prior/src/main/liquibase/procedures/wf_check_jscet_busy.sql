if exists (select '*' from sysprocedure where proc_name like 'wf_jscet_check_busy') then  
	drop function wf_jscet_check_busy;
end if;

create function wf_jscet_check_busy (
	p_numorder integer            // заказ, которому меняем номер счета
	,p_invoice varchar(10)         // новый номер счета заказа
) returns integer
begin


	declare v_check    integer;
	declare v_id_jscet integer;
	declare v_invCode  varchar(10);
	declare v_server   varchar(20);
	declare v_jscet_nu  varchar(10);


	select id_jscet, v.invCode, v.sysname
	into v_id_jscet, v_invCode, v_server
	from orders o
		join guideventure v on v.ventureId = o.ventureId
	where numorder = p_numorder;

	if v_server is null then
		select id_jscet, v.invCode, v.sysname
		into v_id_jscet, v_invCode, v_server
		from bayorders o
			join guideventure v on v.ventureId = o.ventureId
		where numorder = p_numorder;
	end if;

	if v_server is not null then
		set v_jscet_nu = extract_invoice_number(p_invoice, v_invCode);

		call cmt_jscet_check_remote(v_server, v_id_jscet, v_jscet_nu, v_check);
		set wf_jscet_check_busy = v_check;
	else
		set wf_jscet_check_busy = 0;
	end if;
	
end