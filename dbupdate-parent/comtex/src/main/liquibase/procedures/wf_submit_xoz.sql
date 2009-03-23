if exists (select '*' from sysprocedure where proc_name like 'wf_submit_xoz') then  
	drop  procedure wf_submit_xoz;
end if;
 
create procedure wf_submit_xoz (
	  in p_id_accd integer
	, in p_id_accc integer
	, in p_id_m_xoz integer
	, in p_id_deb integer
	, in p_id_jdog integer
	, in p_id      integer
	, in p_dat     date
	, in p_sum     float
	, in p_sumv    float
	, in p_id_curr integer
	, in p_rem    varchar(99)
	, in p_bind_zakaz integer
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
	declare v_id_jscet integer;

	select d.sc, d.sub_sc, d.nm, isnull(d.rem, '')
	into v_debit_sc, v_debit_sub, v_nm, v_rem
	from account d 
	where d.id = p_id_accd;


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
	where c.id = p_id_accc;


	call admin.slave_put_account_prior(
		f_account_exists
		, v_credit_sc
		, v_credit_sub
		, v_nm, v_rem
	);
	if f_account_exists = 0 then
		--return;
	end if;

	if (p_id_m_xoz is not null or p_id_m_xoz != 0) then
		select nm
		into v_purpose
		from m_xoz m
		where m.id = p_id_m_xoz;

		call admin.slave_set_purpose_prior (
	    	  v_purpose
	    	, v_debit_sc, v_debit_sub, v_credit_sc, v_credit_sub 
	    	, v_purpose_id
		);
	end if;

	set v_kredDebitor = admin.wf_kreditor_debitor(p_id_deb);

	select id into v_id_jscet 
	from jscet 
	where id_jdog = p_id_jdog and isnull(id_jdog, 0) != 0;
--	
	select convert(varchar(20), nu) into v_note from jdog where id = p_id_jdog;

	call admin.slave_put_xoz_prior (
		  admin.get_server_name() 
		, p_id
		, v_debit_sc, v_debit_sub, v_credit_sc, v_credit_sub
		, convert(varchar(20), p_dat, 115)
		, p_sum
		, p_sumv
		, p_id_curr
		, p_rem
		, v_id_jscet
		, v_purpose_id
		, v_kredDebitor
		, convert(varchar(20), v_note)
		, p_bind_zakaz
	);

end;
		
