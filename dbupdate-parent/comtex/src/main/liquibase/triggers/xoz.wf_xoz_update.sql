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
	declare v_kredDebitor integer;
--	declare varchar_id varchar(20);
	declare v_values varchar(1024);
	declare v_ventureid char(20);

	declare v_note varchar(10);
	declare v_zakaz varchar(150);
	declare s_id_shiz varchar(20);
	declare v_id_m_xoz integer;
	declare m_xoz_updated integer;
	declare v_already_synced integer;
	declare v_id_jscet integer;

	set m_xoz_updated = 0;
	set v_already_synced = 0;

	set v_ventureid = admin.select_remote('prior', 'guideventure', 'ventureid', 'sysname = ''''' + admin.get_server_name() + '''''');

	if update(dat) then
		call admin.update_remote('prior', 'ybook', 'xDate', '''''' + convert(varchar(20), new_name.dat, 115) + ''''''
			, 'id_xoz=' + convert(varchar(20), old_name.id ) + ' and ventureid = ' + v_ventureid
		);
    		
	end if;


	if update(id_jdog) then

		select nu into v_note from jdog where id = new_name.id_jdog;

		if char_length(v_note) > 0 then
		
			call admin.update_remote('prior', 'ybook', 'note', '''''' + v_note + ''''''
				, 'id_xoz=' + convert(varchar(20), old_name.id ) + ' and ventureid = ' + v_ventureid
			);
    		
			select c.sc
			into v_credit_sc
			from account c 
			where c.id = new_name.id_accc;

			if isnull(new_name.id_jdog, 0) != 0 then
				select id into v_id_jscet from jscet where id_jdog = new_name.id_jdog;
	    
				call admin.slave_bind_zakaz_prior (
					v_zakaz
					, admin.get_server_name()
					, v_note
					, old_name.sum
					, v_credit_sc
					, v_id_jscet
					, old_name.id
				);
			end if;
		end if;

		call admin.wf_synchronize_sum(
			old_name.id_jdog
			, old_name.sum
			, new_name.id_jdog
			, new_name.sum
			, old_name.id
			, v_ventureId
		);
		set v_already_synced = 1;

	end if;

	if update(id_deb) then
		set v_kredDebitor = admin.wf_kreditor_debitor(new_name.id_deb);

		call admin.update_remote('prior', 'ybook', 'KredDebitor', convert(varchar(20), v_kredDebitor)
			, 'id_xoz=' + convert(varchar(20), old_name.id ) + ' and ventureid = ' + v_ventureid
		);

    end if;

	if update(id_accd) and isnull(new_name.id_accd, 0) != 0 then
		    
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

		call admin.update_remote(
			'prior'
			, 'ybook'
			, 'purposeId'
			, 'null'
			, 'id_xoz=' + convert(varchar(20), old_name.id ) + ' and ventureid = ' + v_ventureid
		);

		call admin.update_remote(
			'prior'
			, 'ybook'
			, 'Debit'
			, '''''' + v_debit_sc + ''''''
			, 'id_xoz=' + convert(varchar(20), old_name.id ) + ' and ventureid = ' + v_ventureid
		);

		call admin.update_remote(
			'prior'
			, 'ybook'
			, 'subDebit'
			, '''''' + v_debit_sub + ''''''
			, 'id_xoz=' + convert(varchar(20), old_name.id ) + ' and ventureid = ' + v_ventureid
		);
    
		set m_xoz_updated = 1;
	end if;

	if update(id_accc) and isnull(new_name.id_accc, 0) != 0  then

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

		call admin.update_remote(
			'prior'
			, 'ybook'
			, 'purposeId'
			, 'null'
			, 'id_xoz=' + convert(varchar(20), old_name.id ) + ' and ventureid = ' + v_ventureid
		);

		call admin.update_remote(
			'prior'
			, 'ybook'
			, 'Kredit'
			, ''''''+v_credit_sc+''''''
			, 'id_xoz=' + convert(varchar(20), old_name.id ) + ' and ventureid = ' + v_ventureid
		);

		call admin.update_remote(
			'prior'
			, 'ybook'
			, 'subKredit'
			, ''''''+v_credit_sub+''''''
			, 'id_xoz=' + convert(varchar(20), old_name.id ) + ' and ventureid = ' + v_ventureid
		);
		set m_xoz_updated = 1;

	end if;

	if update(sum) and v_already_synced = 0 then begin
		call admin.wf_synchronize_sum(
			old_name.id_jdog
			, old_name.sum
			, new_name.id_jdog
			, new_name.sum
			, old_name.id
			, v_ventureId
		);

	end; end if;


	if (update(id_m_xoz) or isnull(new_name.id_m_xoz, 0) != 0) 
			or m_xoz_updated = 1 
	then

		select d.sc, d.sub_sc, d.nm, d.rem
		into v_debit_sc, v_debit_sub, v_nm, v_rem
		from account d 
		where d.id = new_name.id_accd;
	    
		select c.sc, c.sub_sc, c.nm, c.rem
		into v_credit_sc, v_credit_sub, v_nm, v_rem
		from account c 
		where c.id = new_name.id_accc;

		call admin.wf_purpose_sync (
			old_name.id
			, v_ventureid
			, new_name.id_m_xoz
			, v_debit_sc
			, v_debit_sub
			, v_credit_sc
			, v_credit_sub
		);
		    

	end if;


	if update(rem) then

		call admin.update_remote(
			'prior'
			, 'ybook'
			, 'descript'
			, ''''''+new_name.rem+''''''
			, 'id_xoz=' + convert(varchar(20), old_name.id ) + ' and ventureid = ' + v_ventureid
		);

	end if;

	if update(id_sh_zatrat) then
		if isnull(new_name.id_sh_zatrat, 0) = 0 then 
			set s_id_shiz = 'null';
		else 
			set s_id_shiz = convert(varchar(20), new_name.id_sh_zatrat);
		end if;

		call admin.update_remote(
			'prior'
			, 'ybook'
			, 'id_shiz'
			, s_id_shiz
			, 'id_xoz=' + convert(varchar(20), old_name.id ) + ' and ventureid = ' + v_ventureid
		);
	end if;

end;

