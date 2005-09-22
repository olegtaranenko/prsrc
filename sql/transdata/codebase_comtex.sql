if exists (select 1 from systriggers where trigname = 'wf_xoz_insert' and tname = 'xoz' and event='INSERT') then 
	drop trigger xoz.wf_xoz_insert;
end if;

create TRIGGER wf_xoz_insert before insert order 211 on
xoz
referencing new as new_name
for each row
--when(new_name.id_guide = 1120 or old_name.id_guide = 1120)
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

	select d.sc, d.sub_sc, d.nm, isnull(d.rem, '')
	into v_debit_sc, v_debit_sub, v_nm, v_rem
	from account d 
	where d.id = new_name.id_accd;


	message 'd.sc = '+v_debit_sc to client;
	message 'd.sub_sc = '+v_debit_sub to client;
	message 'nm = '+v_nm to client;
	message 'rem = '+v_rem to client;

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
	where c.id = new_name.id_accc;

	message 'c.sc = '+v_credit_sc to client;
	message 'c.sub_sc = '+v_credit_sub to client;
	message 'nm = '+v_nm to client;
	message 'rem = '+v_rem to client;


	call admin.slave_put_account_prior(
		f_account_exists
		, v_credit_sc
		, v_credit_sub
		, v_nm, v_rem
	);
	if f_account_exists = 0 then
		--return;
	end if;

	if (new_name.id_m_xoz is not null or new_name.id_m_xoz != 0) then
		select nm
		into v_purpose
		from m_xoz m
		where m.id = new_name.id_m_xoz;

		call admin.slave_set_purpose_prior (
	    	  v_purpose
	    	, v_debit_sc, v_debit_sub, v_credit_sc, v_credit_sub 
	    	, v_purpose_id
		);
	end if;

	call admin.slave_put_xoz_prior(
			  @@servername 
			, new_name.id
			, v_debit_sc, v_debit_sub, v_credit_sc, v_credit_sub
 			, convert(varchar(20), new_name.dat, 115)
			, new_name.sum
			, new_name.sumv
			, new_name.id_curr
			, new_name.rem
			, v_purpose_id
	);
		

end;


if exists (select 1 from systriggers where trigname = 'wf_xoz_update' and tname = 'xoz') then 
	drop trigger xoz.wf_xoz_update;
end if;

create TRIGGER wf_xoz_update before update order 211 on
xoz
referencing old as old_name new as new_name
for each row
begin
	declare v_debit_sc varchar(26);
	declare v_debit_sub varchar(10);
	declare v_credit_sc varchar(26);
	declare v_credit_sub varchar(10);
	declare v_nm varchar(98);
	declare v_rem varchar(98);
	declare f_account_exists integer;
    declare v_currency_rate varchar(20);
    declare v_currency float;
    declare v_m_xoz varchar(100);
    declare v_purposeid integer;

	if update(id_accd) and new_name.id_accd != 0 then
		select d.sc, d.sub_sc, d.nm, d.rem
		into v_debit_sc, v_debit_sub, v_nm, v_rem
		from account d 
		where d.id = new_name.id_accd;

	    
		call admin.slave_put_account_prior(
			f_account_exists
			, v_debit_sc
			, v_debit_sub
			, v_nm, v_rem
		);
		if f_account_exists = 0 then
			--return;
		end if;

		call admin.update_remote(
			'prior'
			, 'ybook y'
			, 'Debit'
			, ''''''+v_debit_sc+''''''
			, 'v.sysname = '''''
					+ @@servername 
					+''''' and v.ventureid = y.ventureid and y.id_xoz = '
					+ convert(varchar(20), old_name.id)
			, 'GuideVenture v'
		);

		call admin.update_remote(
			'prior'
			, 'ybook y'
			, 'subDebit'
			, '''''' + v_debit_sub + ''''''
			, 'v.sysname = '''''
					+ @@servername 
					+''''' and v.ventureid = y.ventureid and y.id_xoz = '
					+ convert(varchar(20), old_name.id)
			, 'GuideVenture v'
		);
	end if;

	if update(id_accc) and new_name.id_accc != 0 then
		select c.sc, c.sub_sc, c.nm, c.rem
		into v_credit_sc, v_credit_sub, v_nm, v_rem
		from account c 
		where c.id = new_name.id_accc;
	    
		call admin.slave_put_account_prior(
			f_account_exists
			, v_credit_sc
			, v_credit_sub
			, v_nm, v_rem
		);
		if f_account_exists = 0 then
			--return;
		end if;

		call admin.update_remote(
			'prior'
			, 'ybook y'
			, 'Kredit'
			, ''''''+v_credit_sc+''''''
			, 'v.sysname = '''''
					+ @@servername 
					+''''' and v.ventureid = y.ventureid and y.id_xoz = '
					+ convert(varchar(20), old_name.id)
			, 'GuideVenture v'
		);

		call admin.update_remote(
			'prior'
			, 'ybook y'
			, 'subKredit'
			, ''''''+v_credit_sub+''''''
			, 'v.sysname = '''''
					+ @@servername 
					+''''' and v.ventureid = y.ventureid and y.id_xoz = '
					+ convert(varchar(20), old_name.id)
			, 'GuideVenture v'
		);
	end if;

	if update(sum) then
		call admin.slave_select_prior(
			v_currency_rate
			,'system'
			,'Kurs'
			,'1=1'
		);

		set v_currency = new_name.sum / convert(float, abs(v_currency_rate));

		call admin.update_remote(
			'prior'
			, 'ybook y'
			, 'UEsumm'
			, v_currency
			, 'v.sysname = '''''
					+ @@servername 
					+''''' and v.ventureid = y.ventureid and y.id_xoz = '
					+ convert(varchar(20), old_name.id)
			, 'GuideVenture v'
		);

	end if;


	if update(id_m_xoz) and new_name.id_m_xoz!= 0 then
		select d.sc, d.sub_sc, d.nm, d.rem
		into v_debit_sc, v_debit_sub, v_nm, v_rem
		from account d 
		where d.id = old_name.id_accd;
	    
		select c.sc, c.sub_sc, c.nm, c.rem
		into v_credit_sc, v_credit_sub, v_nm, v_rem
		from account c 
		where c.id = old_name.id_accc;

		select nm into v_m_xoz from m_xoz where id = new_name.id_m_xoz;

		if v_debit_sc is not null and v_credit_sc is not null then
			call admin.slave_set_purpose_prior(
				 v_m_xoz
				, v_debit_sc, v_debit_sub, v_credit_sc, v_credit_sub
				, v_purposeid
			);
			//message 'v_purposeid = ' + convert(varchar(20), v_purposeid) to client;	
			call admin.update_remote(
				'prior'
				, 'ybook y'
				, 'purposeId'
				, v_purposeid
				, 'v.sysname = '''''
						+ @@servername 
						+''''' and v.ventureid = y.ventureid and y.id_xoz = '
						+ convert(varchar(20), old_name.id)
				, 'GuideVenture v'
			);
		end if;

	end if;
	if update(rem) then
		select d.sc, d.sub_sc, d.nm, d.rem
		into v_debit_sc, v_debit_sub, v_nm, v_rem
		from account d 
		where d.id = old_name.id_accd;
	    
		select c.sc, c.sub_sc, c.nm, c.rem
		into v_credit_sc, v_credit_sub, v_nm, v_rem
		from account c 
		where c.id = old_name.id_accc;

		call admin.update_remote(
			'prior'
			, 'ybook y'
			, 'descript'
			, ''''''+new_name.rem+''''''
			, 'v.sysname = '''''
					+ @@servername 
					+''''' and v.ventureid = y.ventureid and y.id_xoz = '
					+ convert(varchar(20), old_name.id)
			, 'GuideVenture v'
		);

	end if;
end;


if exists (select 1 from systriggers where trigname = 'wf_xoz_delete' and tname = 'xoz') then 
	drop trigger xoz.wf_xoz_delete;
end if;

create TRIGGER wf_xoz_delete before delete order 211 on
xoz
referencing old as old_name
for each row
begin
	call admin.slave_delete_prior(
		 'ybook y'
		, 'v.sysname = '''
				+ @@servername 
				+ ''' and v.ventureid = y.ventureid and y.id_xoz = '
				+ convert(varchar(20), old_name.id)
		, 'GuideVenture v'
	);
end;

