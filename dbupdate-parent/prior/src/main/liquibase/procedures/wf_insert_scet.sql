if exists (select '*' from sysprocedure where proc_name like 'wf_insert_scet') then  
	drop procedure wf_insert_scet;
end if;

create function wf_insert_scet (
	  p_servername varchar(20)
	, p_id_jscet integer
	, p_id_inv integer
	, p_quant double
	, p_cena double
	, p_date date
	, in p_rate double
	, in p_ndsrate double
)
returns integer
begin
	declare v_id_scet integer;
	declare v_fields varchar(255);
	declare v_values varchar(2000);
	declare scet_nu integer;
--	declare v_currency_rate double;
	declare v_datev varchar(20);
	declare v_id_cur integer;


  set wf_insert_scet =  null;
 
  if p_servername is not null and p_id_jscet is not null then

	-- Получить следующий порядковый номер счета бух.базы
	set scet_nu = select_remote(
		p_servername
		, 'scet'
		, 'max(nu)+1'
		, 'id_jmat = ' + convert(varchar(20), p_id_jscet)
	);

	set scet_nu = isnull(scet_nu, 1);

	-- По какому курсу, учитывая, что в бухгалтерии только рубли, а в Приоре - УЕ
	set v_id_cur = system_currency();

	
	set v_fields = '
		 id_jmat
		,id_inv
		,kol1
		,nu
		,summa_sale
		,summa_salev
		,summa_nds
	';

	set v_values = 
		convert(varchar(20), p_id_jscet)
		+', '+ convert(varchar(20), p_id_inv)
		+', '+ convert(varchar(20), p_quant)
		+', '+ convert(varchar(20), scet_nu)
		+', '+ convert(varchar(20), round(p_quant*p_cena * p_rate, 2))
		+', '+ convert(varchar(20), round(p_quant*p_cena, 2))
		+', '+ convert(varchar(20), round(p_quant*p_cena*p_ndsrate/100, 2))
	;
	--message 'p_cena = ', p_cena to client;
	--message 'p_quant = ', p_quant to client;
	--message 'v_values = ', v_values to client;

	-- изменения в бухгалтерской базе данных
	set v_id_scet = insert_count_remote(p_servername, 'scet', v_fields, v_values);

	set wf_insert_scet = v_id_scet;
  end if;

end;

