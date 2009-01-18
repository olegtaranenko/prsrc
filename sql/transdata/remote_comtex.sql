if not exists (select 1 from sys.sysservers where srvname = 'prior') then  
	create server prior class 'ASAODBC' USING 'DSN=prior;UID=dba;PWD=sql';
end if;



call build_host_procedure ( 
		 'put_account'
		, '  out out_exists integer'
		+ ', inout p_sc char(26)'
		+ ', inout p_sub char(10)'
		+ ', inout p_name char(98)'
		+ ', inout p_desc char(98)'
);

call build_host_procedure (
		  'bind_zakaz', 
		  '  out v_orderNum char(150)'
		+ ', p_server     char(50)'
		+ ', p_invoice	  char(10)'
		+ ', in p_sum        real'
		+ ', in p_sc_credit char(10)'
		+ ', in p_id_xoz  integer default null'
);



call build_host_procedure (
		  'set_purpose', 
		  '  p_purpose     char(99)'
		+ ', p_debit_sc    char(26)'
		+ ', p_debit_sub   char(10)'
		+ ', p_credit_sc   char(26)'
		+ ', p_credit_sub  char(10)'
		+ ', out p_purposeId integer'
);


