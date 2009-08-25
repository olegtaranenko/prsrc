if exists (select 1 from systriggers where trigname = 'wf_update_orders' and tname = 'BayOrders') then 
	drop trigger BayOrders.wf_update_orders;
end if;

create TRIGGER wf_update_orders before update on
BayOrders
referencing old as old_name new as new_name
for each row
begin
	declare remoteServerOld varchar(32);
	declare remoteServerNew varchar(32);
	declare v_fields varchar(255);
	declare v_values varchar(2000);
	declare v_nu_jscet integer;
	declare r_nu varchar(50);
	declare r_id integer;
--	declare v_firm_id integer;
	declare v_invCode varchar(10);
	declare v_id_dest integer;
	declare v_id_schef integer;
	declare v_id_bux integer;
	declare v_id_bank integer;
	declare v_datev varchar(20);
	declare v_id_cur integer;
	declare v_inv_date varchar(20);
	declare v_numOrder integer;
	declare sync char(1);
	declare c_status_close_id integer;
	declare v_ivo_procent float;
	declare v_updated integer;
	declare v_nu_jdog varchar(17);
	declare v_id_jdog integer;


	set c_status_close_id = 6;  -- закрыт
	select sysname, invCode into remoteServerOld, v_invcode from GuideVenture where ventureId = old_name.ventureId;

	if update(invoice) and remoteServerOld is not null then begin

		set v_nu_jscet = extract_invoice_number(new_name.invoice, v_invCode);
		set v_id_jdog = select_remote(remoteServerOld, 'jscet', 'id_jdog', 'id = ' + convert(varchar(20), old_name.id_jscet));
		set v_nu_jdog = wf_make_jdog_nu(v_nu_jscet, old_name.inDate);

		call block_remote(remoteServerOld, get_server_name(), 'jscet');


		call update_remote(remoteServerOld, 'jdog', 'nu',  '''''' + v_nu_jdog  + '''''', 'id = ' + convert(varchar(20), v_id_jdog));
		call unblock_remote(remoteServerOld, get_server_name(), 'jscet');

		call update_remote(remoteServerOld, 'jscet', 'nu'
				, convert(varchar(20), v_nu_jscet)
				, 'id = ' + convert(varchar(20), old_name.id_jscet)
		);

	end; end if;

	if update(ventureId) then
		if new_name.ventureId = 0 then
			set new_name.ventureid = null;
		end if;
		if isnull(old_name.ventureId, 0) != isnull(new_name.ventureId, 0) then
			if remoteServerOld is not null then
				call purge_jscet(remoteServerOld, old_name.id_jscet);
				set new_name.invoice = 'счет ?';
				set new_name.id_bill = null;
			end if;

			select sysname, invCode into remoteServerNew, v_invcode from GuideVenture where ventureId = new_name.ventureId;

			--message 'sysname = ', remoteServerNew to client;
			if remoteServerNew is not null then
	
				set v_numOrder = old_name.numOrder;
--				set v_firm_id = old_name.firmId;
				select id_voc_names into v_id_dest from bayguidefirms where firmid = old_name.firmId;
				call put_jscet(r_id, v_nu_jscet, remoteServerNew, v_numOrder, v_id_dest, old_name.invoice, old_name.rate);
		
				set new_name.id_jscet = r_id;
				set new_name.invoice = v_invCode + convert(varchar(20), v_nu_jscet);
				call wf_set_bay_detail(remoteServerNew, r_id, new_name.numOrder, v_inv_date, old_name.rate);
			end if;
		end if;
	end if;
	if update (firmId) then
		if remoteServerOld is not null then
			select id_voc_names into v_id_dest from BayGuideFirms where firmId = new_name.firmId;
			call block_remote(remoteServerOld, get_server_name(), 'jscet');
			call update_remote(remoteServerOld, 'jscet', 'id_d', convert(varchar(20), v_id_dest ), 'id = ' + convert(varchar(20), old_name.id_jscet));
			call update_remote(remoteServerOld, 'jscet', 'id_d_cargo', convert(varchar(20), v_id_dest ), 'id = ' + convert(varchar(20), old_name.id_jscet));
			call unblock_remote(remoteServerOld, get_server_name(), 'jscet');
		end if;
	end if;
	if 	update (statusId)
		and new_name.statusId = c_status_close_id
--		and wf_order_closed_comtex(old_name.numorder, remoteServerOld) = 1 
	then
		select ivo_procent into v_ivo_procent from system;
--		set v_numorder = old_name.numorder;
		-- генерить взаимозачеты
		call ivo_generate_numdoc(old_name.numorder, v_ivo_procent);

	end if;

	if update(rate) then
		if remoteServerOld is not null then
			select id_cur into v_id_cur from system;
			call update_remote(remoteServerOld, 'jscet', 'curr', convert(varchar(20), new_name.rate ), 'id = ' + convert(varchar(20), old_name.id_jscet));
			call update_remote(remoteServerOld, 'jscet', 'id_curr', convert(varchar(20), v_id_cur ), 'id = ' + convert(varchar(20), old_name.id_jscet));
			for x as xxc dynamic scroll cursor for
				select r.id_scet as r_id_scet
					, r.intQuant / n.perList as r_cenaEd
					, r.quantity as r_quant
				from sdmcrez r 
				join sGuideNomenk n on n.nomnom = r.nomnom
				where numdoc = old_name.numorder
			do
				set v_updated = wf_scet_price_changed(remoteServerOld, r_quant, r_cenaEd, r_id_scet, new_name.rate);
			end for;
		end if;

	end if;

end;


