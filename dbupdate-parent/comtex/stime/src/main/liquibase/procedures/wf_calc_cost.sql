if exists (select '*' from sysprocedure where proc_name like 'wf_calc_cost') then  
	drop procedure wf_calc_cost;
end if;


create procedure wf_calc_cost (
	  out out_ret        float
	, out out_has_naklad integer
	, p_id_inv           integer
) 
begin

	--execute immediate 'create variable @adec_Ost21 decimal(19, 7)';

	--call calc_ost_inv(now(), p_id_inv, -1, -2,  '1' , '2' , '1' , 0 , '0' , '0' , 1 , 1 , '0' , '0' , '0' , 0 );
	set out_ret = calc_summa('mat', -1, now(), p_id_inv, -2, 'summa', 1, 7);

	--message summa_rub, ' ', @adec_Ost21 to client;

--	set v_string_prc = select_remote('stime', 'inv', 'prc1', 'id = ' + convert(varchar(20), p_id_inv));
	--set out_ret = summa_rub / @adec_Ost21;
	--execute immediate 'drop variable @adec_Ost21';

	-- flag shows whether the item has movement or not
	select count(*) into out_has_naklad from scet where id_inv = p_id_inv;
end;


