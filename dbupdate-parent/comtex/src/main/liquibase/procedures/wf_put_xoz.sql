if exists (select '*' from sysprocedure where proc_name like 'wf_put_xoz') then  
	drop procedure wf_put_xoz;
end if;

create procedure wf_put_xoz (
	  new_name_id         integer
	, new_name_dat        timestamp
	, new_name_sum        float
	, new_name_sumv       float
	, new_name_id_curr    integer
	, new_name_id_accd    integer
	, new_name_id_accc    integer
	, new_name_id_deb     integer
	, new_name_id_jdog    integer
	, new_name_id_m_xoz   integer
	, new_name_rem        varchar(99)
) 
begin
	declare v_debit_sc varchar(26);
	declare v_debit_sub varchar(10);
	declare v_credit_sc varchar(26);
	declare v_credit_sub varchar(10);
	declare f_account_exists integer;
	declare v_nm varchar(98);
	declare v_rem varchar(98);
	declare v_purpose_id integer;
	declare v_detail_id integer;
	declare v_purpose varchar(99);
	declare v_kredDebitor integer;
	declare v_note varchar(50);

	select d.sc, d.sub_sc, d.nm, isnull(d.rem, '')
	into v_debit_sc, v_debit_sub, v_nm, v_rem
	from account d 
	where d.id = new_name_id_accd;


	call admin.slave_put_account_prior(
		f_account_exists
		, v_debit_sc
		, v_debit_sub
		, v_nm
		, v_rem
	);
	if f_account_exists = 0 then
		--return;
	end if;

	select c.sc, c.sub_sc, c.nm, c.rem
	into v_credit_sc, v_credit_sub, v_nm, v_rem
	from account c 
	where c.id = new_name_id_accc;


	call admin.slave_put_account_prior(
		f_account_exists
		, v_credit_sc
		, v_credit_sub
		, v_nm, v_rem
	);
	if f_account_exists = 0 then
		--return;
	end if;

	if (new_name_id_m_xoz is not null or new_name_id_m_xoz != 0) then
		select nm
		into v_purpose
		from m_xoz m
		where m.id = new_name_id_m_xoz;

		call admin.slave_set_purpose_prior (
	    	  v_purpose
	    	, v_debit_sc, v_debit_sub, v_credit_sc, v_credit_sub 
	    	, v_purpose_id
		);
	end if;

	set v_kredDebitor = admin.wf_kreditor_debitor(new_name_id_deb);

	select nu into v_note from jdog where id = new_name_id_jdog;

	call admin.slave_put_xoz_prior(
		  admin.get_server_name() 
		, new_name_id
		, v_debit_sc, v_debit_sub, v_credit_sc, v_credit_sub
		, convert(varchar(20), new_name_dat, 115)
		, new_name_sum
		, new_name_sumv
		, new_name_id_curr
		, new_name_rem
		, v_purpose_id
		, v_kredDebitor
		, convert(varchar(20), v_note)
	);
		
end;

