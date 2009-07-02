if exists (select 1 from systriggers where trigname = 'wf_update_orders' and tname = 'Orders') then 
	drop trigger Orders.wf_update_orders;
end if;

create TRIGGER wf_update_orders before update order 1 on
Orders
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
	declare v_currency_rate float;
	declare v_cenaEd float;
	declare v_order_date varchar(20);
	declare v_check_count integer; 
	declare v_id_jscet integer;
	declare v_id_scet integer;
	declare v_id_inv integer;
	declare v_numorder integer;
	declare v_updated integer;

	declare v_issue_id integer;
	declare v_msgCode integer;

--	declare v_total_account_date datetime;
	declare sync char(1);
	declare c_status_close_id integer;
	declare v_ivo_procent float;

	set c_status_close_id = 6;  -- закрыт

	select sysname, invCode into remoteServerOld, v_invcode from GuideVenture where ventureId = old_name.ventureId;
	set v_currency_rate = new_name.rate;
	select id_cur into v_id_cur from system;
	select sysname into remoteServerOld from GuideVenture where ventureId = old_name.ventureId;

	if update(invoice) and remoteServerOld is not null then begin
		declare v_nu_jdog varchar(17);
		declare v_id_jdog integer;

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
			if remoteServerNew is not null then
		
				set v_numOrder = old_name.numOrder;
				select id_voc_names into v_id_dest from guidefirms where firmid = old_name.firmId;
--				set v_firm_id = old_name.firmId;
				call put_jscet(r_id, v_nu_jscet, remoteServerNew, v_numOrder, v_id_dest, old_name.invoice, v_currency_rate);
		
				set new_name.id_jscet = r_id;
				set new_name.invoice = v_invCode + convert(varchar(20), v_nu_jscet);
				call wf_set_invoice_detail(remoteServerNew, r_id, new_name.numOrder, v_order_date, v_currency_rate);
			end if;

			-- исправление расходных накладных, связанных с заказом
			--select total_account into v_total_account_date from system;

			-- это можно делать только для тех заказов, которые после перехода на режим полного учета по предприятиям
			update sdocs set ventureId = new_name.ventureId 
			where 
				sdocs.numdoc = new_name.numorder
				--and xDate >= v_total_account_date
			;
		end if;

	end if;

	
	
	if update (ordered) or update(rate) then

		set v_id_jscet = old_name.id_jscet;
	
		if remoteServerOld is not null and v_id_jscet is not null then

			if update(rate) then
				call update_remote(remoteServerOld, 'jscet', 'curr', convert(varchar(20), v_currency_rate ), 'id = ' + convert(varchar(20), old_name.id_jscet));
				call update_remote(remoteServerOld, 'jscet', 'id_curr', convert(varchar(20), v_id_cur ), 'id = ' + convert(varchar(20), old_name.id_jscet));
			end if;

			
			-- Заказ, который имеет ссылки в бух.базах интеграции
			-- т.е. уже назначен той, иди другой фирме

			-- отследить заказ без предметов
			-- сначала проверяем, что он действительно без них
			set v_check_count = 0;
			for x as xc dynamic scroll cursor for
				select n.id_scet as r_id_scet, n.cenaEd as r_cenaEd, n.quant as r_quant, n.nomnom as r_nomnom
					, k.ed_izmer as r_edizm, trim(k.cod + ' ' + k.nomname + ' ' + k.size) as r_nomenk
				from xpredmetybynomenk n
				join sguidenomenk k on k.nomnom = n.nomnom
				where numorder = old_name.numorder
			do
				set v_check_count = 1;
				set v_updated = wf_scet_price_changed(remoteServerOld, r_quant, r_cenaEd, r_id_scet, v_currency_rate);
				if v_updated = 0 then
					-- Этой позиции нет в комтехе. Вероятно бухгалтер изменил номенклатуру
					if v_issue_id is null then
						set v_issue_id = wi_check_business_issue(@issueMarker);
						if v_issue_id is null then
							set v_issue_id = wi_post_new_issue('ref-nomenk-missed', 'prior');
						end if;
						call wi_add_issue_attribute(v_issue_id, 'Номер заказа', old_name.numorder);
						call wi_add_issue_attribute(v_issue_id, 'Номер счета', old_name.invoice);
						call wi_add_issue_attribute(v_issue_id, 'Курс заказа', old_name.rate);
						call wi_add_issue_attribute(v_issue_id, 'Новый курс', new_name.rate);
					end if;
					call wi_add_issue_attribute(v_issue_id, 'Номенклатура', r_nomenk);
					call wi_add_issue_attribute(v_issue_id, 'Номер номенклатуры', r_nomnom);
					call wi_add_issue_attribute(v_issue_id, 'Количество', r_quant);
					call wi_add_issue_attribute(v_issue_id, 'Единица измерения', r_edizm);
					call wi_add_issue_attribute(v_issue_id, 'Цена за единицу', r_cenaEd);
				end if;
			end for;

			for y as yc dynamic scroll cursor for
				select i.id_scet as r_id_scet, i.cenaEd as r_cenaEd, i.quant as r_quant, i.prId as r_prId
					, trim (p.prName + ' ' + p.prDescript + ' ' + p.prSize) as r_productDescript
				from xpredmetybyizdelia i
				join sguideProducts p on p.prId = i.prId
				where numorder = old_name.numorder
			do
				set v_check_count = 1;
				set v_updated = wf_scet_price_changed(remoteServerOld, r_quant, r_cenaEd, r_id_scet, v_currency_rate);
				if v_updated = 0 then
					-- Этой позиции нет в комтехе. Вероятно бухгалтер изменил номенклатуру
					if v_issue_id is null then
						set v_issue_id = wi_check_business_issue(@issueMarker);
						if v_issue_id is null then
							set v_issue_id = wi_post_new_issue('ref-nomenk-missed', 'prior');
						end if;
						call wi_add_issue_attribute(v_issue_id, 'Номер заказа', old_name.numorder);
						call wi_add_issue_attribute(v_issue_id, 'Номер счета', old_name.invoice);
						call wi_add_issue_attribute(v_issue_id, 'Курс заказа', old_name.rate);
						call wi_add_issue_attribute(v_issue_id, 'Новый курс', new_name.rate);
					end if;
					call wi_add_issue_attribute(v_issue_id, 'Изделие', r_productDesc);
					call wi_add_issue_attribute(v_issue_id, 'ID изделия', r_prID);
					call wi_add_issue_attribute(v_issue_id, 'Индекс изделия в заказе', r_prExt);
					call wi_add_issue_attribute(v_issue_id, 'Количество', r_quant);
					call wi_add_issue_attribute(v_issue_id, 'Цена за единицу', r_cenaEd);
				end if;
			end for;

	    
			if v_check_count > 0 then
				-- заказ с предметами
				if v_issue_id is not null then
					-- номенклатура в бухгалтерии не соответствует в приоре.
					set v_msgCode = wi_get_msgcode(v_issue_id);
					--raiserror v_msgCode '{{{' + convert(varchar(20), v_issue_id) + '}}}';
				end if;

				return;
			end if;
	    
			-- ищем товар под названием "услуга"
			select id_inv into v_id_inv from sGuideNomenk where nomNom = 'УСЛ';

			-- сначала исходим из того, что такая услуга уже есть.
			-- это может произойти при изменении стоимости заказа.

			if abs(new_name.ordered) < 0.001 then
				call delete_remote(remoteServerOld, 'scet'
					, 'id_jmat = ' + convert(varchar(20), v_id_jscet) + ' and id_inv = ' + convert(varchar(20), v_id_inv)
				);
				return;
			end if;

			set v_id_scet = null;
			set v_id_scet = select_remote(remoteServerOld, 'scet', 'id', 'id_jmat = ' + convert(varchar(20), v_id_jscet) + ' and id_inv = ' + convert(varchar(20), v_id_inv));
			set v_cenaEd = 1; -- цена услуги - 1 УЕ
			if v_id_scet is not null then

				-- именно такой случай
				set v_updated = wf_scet_price_changed(remoteServerOld, new_name.ordered, v_cenaEd, v_id_scet, v_currency_rate);

			else
				-- первый раз меням это поле => нужно добавить
				set v_id_scet = 
					wf_insert_scet(
						remoteServerOld
						, v_id_jscet
						, v_id_inv
						, v_cenaEd
						, new_name.ordered
						, old_name.indate
						, v_currency_rate
					);
			end if;
		end if;


	end if;
	
	
	if update (firmId) and (old_name.id_bill is null or old_name.id_bill = 0) then
		
		if remoteServerOld is not null then
			select id_voc_names into v_id_dest from guideFirms where firmId = new_name.firmId;
			call block_remote(remoteServerOld, get_server_name(), 'jscet');
			call update_remote(remoteServerOld, 'jscet', 'id_d', convert(varchar(20), v_id_dest ), 'id = ' + convert(varchar(20), old_name.id_jscet));
			call update_remote(remoteServerOld, 'jscet', 'id_d_cargo', convert(varchar(20), v_id_dest ), 'id = ' + convert(varchar(20), old_name.id_jscet));
			call unblock_remote(remoteServerOld, get_server_name(), 'jscet');
		end if;
	end if;


	if update (statusId)
		and new_name.statusId = c_status_close_id
--		and wf_order_closed_comtex(old_name.numorder, remoteServerOld) = 1 
	then
		select ivo_procent into v_ivo_procent from system;
--		set v_numorder = old_name.numorder;
		-- генерить взаимозачеты
		call ivo_generate_numdoc(old_name.numorder, v_ivo_procent);

	end if;
/*
	if update (id_bill)  then
		begin
			declare v_nu_jdog varchar(17);
			declare v_id_jdog integer;

			set v_nu_jscet = select_remote(remoteServerOld, 'jscet', 'nu', 'id = ' + convert(varchar(20), old_name.id_jscet));
			call block_remote(remoteServerOld, get_server_name(), 'jscet');

			set v_nu_jdog = wf_make_jdog_nu(v_nu_jscet, old_name.inDate);
			set v_id_jdog = select_remote(remoteServerOld, 'jdog', 'id', 'nu = ''''' + convert(varchar(20), v_nu_jdog) + '''''');

			call update_remote(remoteServerOld, 'jscet', 'id_jdog', convert(varchar(20), v_id_jdog ), 'id = ' + convert(varchar(20), old_name.id_jscet));
			call unblock_remote(remoteServerOld, get_server_name(), 'jscet');
		end;
	end if;
*/
end;
