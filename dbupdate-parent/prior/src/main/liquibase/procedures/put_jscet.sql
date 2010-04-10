if exists (select '*' from sysprocedure where proc_name like 'put_jscet') then  
	drop procedure put_jscet;
end if;

create procedure put_jscet (
	  out r_id integer
	, out v_nu_jscet varchar(50)
	, in remoteServerNew varchar(20)
	, in p_numOrder integer
	, in p_id_dest integer
	, in p_nu_old varchar(50) default null 
	, in p_rate float
) 
begin
	declare v_fields varchar(255);
	declare v_values varchar(2000);
	declare r_nu varchar(50);
--	declare v_firm_id integer;
	declare v_invCode varchar(10);
--	declare p_id_dest integer;
	declare v_id_schef integer;
	declare v_id_bux integer;
	declare v_id_bank integer;
	declare v_datev varchar(20);
	declare v_id_cur integer;
--	declare v_currency_rate float;
	declare v_order_date varchar(20);
	declare v_check_count integer; 
	declare v_id_jscet integer;
	declare v_intInvoice integer;

	declare v_id_jdog integer;
	declare v_jdog_nu varchar(17);
	declare v_exists integer;


	select invCode into v_invCode
	from guideVenture where sysname = remoteServerNew;


	set v_nu_jscet = nextnu_remote(remoteServerNew, 'jscet', p_nu_old);

	set r_id = r_id + 1;
	set v_order_date = convert(varchar(20), now());
	set v_id_cur = system_currency();
	execute immediate 'call slave_currency_rate_' + remoteServerNew + '(v_datev, v_currency_rate, v_order_date, v_id_cur )';

	set v_fields =
		 'nu'
--		+ ', id'
		+ ', rem'
		+ ', id_s'
		+ ', dat' 
		+ ', datv' 
		+ ', state'
		+ ', real_days'
		+ ', id_curr'
		+ ', curr'
//		+ ', id_kad_bux'
//		+ ', id_s_bank'
		;
	set v_values = 
		convert(varchar(20), v_nu_jscet)
--		+ ', ' + convert(varchar(20), r_id)
		+ ', ' + convert(varchar(20), p_numOrder)
		+ ', -1'
		+ ', ''''' + convert(varchar(20), v_order_date, 112) + ''''''
		+ ', ''''' + v_datev + ''''''
		+ ', 1'
		+ ', 3'
		+ ', ' + convert(varchar(20), v_id_cur)
		+ ', ' + convert(varchar(20), p_rate)
	;


	if p_id_dest is not null then

		set v_exists = select_remote(remoteServerNew, 'voc_names', 'count(*)', 'id = ' + convert(varchar(20), p_id_dest));
    
		if v_exists = 1 then
			set v_fields = v_fields
				+ ', id_d'
				+ ', id_d_cargo'
			;
			set v_values = v_values	
				+ ', ' + convert(varchar(20), p_id_dest)
				+ ', ' + convert(varchar(20), p_id_dest)
			;
    
			-- теперь автоматом генерится договор для данного счета
			-- номер договора имеет шаблоy уууу/ннн, где уууу год, ннн - номер только что добавленного счета
    
			call put_jdog(v_id_jdog, v_jdog_nu, remoteServerNew, v_nu_jscet, p_id_dest, now());
			
			if v_id_jdog is not null then
				set v_fields = v_fields
					+ ', id_jdog'
				;
				set v_values = v_values	
				+ ', ' + convert(varchar(20), v_id_jdog)
				;
			end if;
		end if;
	end if;
	
	set r_id = insert_count_remote(remoteServerNew, 'jscet', v_fields, v_values);


end;

