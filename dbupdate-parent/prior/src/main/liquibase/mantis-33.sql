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
--	declare v_total_account_date datetime;
	declare sync char(1);
	declare c_status_close_id integer;
	declare v_ivo_procent float;

	set c_status_close_id = 6;  -- закрыт

	select sysname, invCode into remoteServerOld, v_invcode from GuideVenture where ventureId = old_name.ventureId;

	if update(invoice) and remoteServerOld is not null then
		call update_remote(remoteServerOld, 'jscet', 'nu'
				, convert(varchar(20), extract_invoice_number(new_name.invoice, v_invCode))
				, 'id = ' + convert(varchar(20), old_name.id_jscet)
		);
	end if;


	if update(ventureId) then
		if new_name.ventureId = 0 then
			set new_name.ventureid = null;
		end if;
		if isnull(old_name.ventureId, 0) != isnull(new_name.ventureId, 0) then
			if remoteServerOld is not null then
				call delete_remote(remoteServerOld, 'jscet', 'id = ' + convert(varchar(20), old_name.id_jscet));
				call delete_remote(remoteServerOld, 'scet', 'id_jmat = ' + convert(varchar(20), old_name.id_jscet));
				set new_name.invoice = 'счет ?';
				set new_name.id_bill = null;
			end if;
		
			select sysname, invCode into remoteServerNew, v_invcode from GuideVenture where ventureId = new_name.ventureId;
			if remoteServerNew is not null then
		
				set v_numOrder = old_name.numOrder;
				select id_voc_names into v_id_dest from guidefirms where firmid = old_name.firmId;
--				set v_firm_id = old_name.firmId;
				call put_jscet(r_id, v_nu_jscet, remoteServerNew, v_numOrder, v_id_dest, old_name.invoice);
		
				set new_name.id_jscet = r_id;
				set new_name.invoice = v_invCode + convert(varchar(20), v_nu_jscet);
				call wf_set_invoice_detail(remoteServerNew, r_id, new_name.numOrder, v_order_date);
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

		set v_currency_rate = new_name.rate;
		set v_id_jscet = old_name.id_jscet;
	
		if remoteServerOld is not null and v_id_jscet is not null then

			if update(rate) then
				call update_remote(remoteServerOld, 'jscet', 'curr', convert(varchar(20), v_currency_rate ), 'id = ' + convert(varchar(20), old_name.id_jscet));
			end if;

			
			-- Заказ, который имеет ссылки в бух.базах интеграции
			-- т.е. уже назначен той, иди другой фирме

			-- отследить заказ без предметов
			-- сначала проверяем, что он действительно без них
			set v_check_count = 0;
			for x as xc dynamic scroll cursor for
				select id_scet as r_id_scet, cenaEd as r_cenaEd, quant as r_quant
				from xpredmetybynomenk 
				where numorder = old_name.numorder
			do
				set v_check_count = 1;
				set v_updated = wf_scet_price_changed(remoteServerOld, r_quant, r_cenaEd, r_id_scet, v_currency_rate);
			end for;

			for y as yc dynamic scroll cursor for
				select id_scet as r_id_scet, cenaEd as r_cenaEd, quant as r_quant
				from xpredmetybyizdelia
				where numorder = old_name.numorder
			do
				set v_check_count = 1;
				set v_updated = wf_scet_price_changed(remoteServerOld, r_quant, r_cenaEd, r_id_scet, v_currency_rate);
			end for;

	    
			if v_check_count > 0 then
				-- заказ с предметами
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
					);
			end if;
		end if;


	end if;
	
	
	
	
	
	
	
	
	if update (firmId) and (old_name.id_bill is null or old_name.id_bill = 0) then
		
		select sysname into remoteServerOld from GuideVenture where ventureId = old_name.ventureId;
		if remoteServerOld is not null then
			select id_voc_names into v_id_dest from guideFirms where firmId = new_name.firmId;
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

end;


if exists (select 1 from systriggers where trigname = 'wf_update_izd' and tname = 'xPredmetyByIzdelia') then 
	drop trigger xPredmetyByIzdelia.wf_update_izd;
end if;

create TRIGGER wf_update_izd before update on
xPredmetyByIzdelia
referencing old as old_name new as new_name
for each row
begin
	declare v_id_scet integer;
	declare v_numorder integer;
	declare v_belong_id integer;
	declare remoteServerNew varchar(32);
	declare v_values varchar(100);
	declare v_fields varchar(200);
	declare v_currency_rate float;
	
	set v_id_scet = old_name.id_scet;
--	set v_numorder = old_name.numOrder;

	select sysname, rate
	into remoteServerNew, v_currency_rate
		from orders o
		join GuideVenture v on o.ventureId = v.ventureId and v.standalone = 0
		where numOrder = old_name.numOrder;


	if remoteServerNew is not null then
		if update(quant) or update(cenaEd) then
			call wf_scet_price_changed(remoteServerNew, new_name.quant, new_name.cenaEd, v_id_scet, v_currency_rate)
		end if;
		if update(quant) then
			call update_remote(remoteServerNew, 'scet', 'kol1', convert(varchar(20), new_name.quant), 'id = ' + convert(varchar(20), v_id_scet));
		end if;
	end if;
  
end;



if exists (select 1 from systriggers where trigname = 'wf_update_nomenk' and tname = 'xPredmetyByNomenk') then 
	drop trigger xPredmetyByNomenk.wf_update_nomenk;
end if;

create TRIGGER wf_update_nomenk before update on
xPredmetyByNomenk
referencing old as old_name new as new_name
for each row
begin
	declare v_id_scet integer;
	declare v_belong_id integer;
	declare remoteServerNew varchar(32);
	declare v_currency_rate float;

	declare v_values varchar(100);
	declare v_fields varchar(200);
	
	set v_id_scet = old_name.id_scet;

	select sysname, rate
	into remoteServerNew, v_currency_rate
		from orders o
		join GuideVenture v on o.ventureId = v.ventureId and v.standalone = 0
		where numOrder = old_name.numOrder;


	if remoteServerNew is not null then
		if update(quant) or update(cenaEd) then
			call wf_scet_price_changed(remoteServerNew, new_name.quant, new_name.cenaEd, v_id_scet, v_currency_rate)
        end if;
		if update(quant) then
			call update_remote(remoteServerNew, 'scet', 'kol1', convert(varchar(20), new_name.quant), 'id = ' + convert(varchar(20), v_id_scet));
		end if;
	end if;
	  
end;
	

if exists (select 1 from sysprocedure where proc_name = 'wf_scet_price_changed') then
	drop function wf_scet_price_changed;
end if;


CREATE function wf_scet_price_changed (
-- апдейтим цены(руб и валютную) в бух базе комтеха при изменении суммы в приоре.
	  p_server_new    varchar(32)
	, p_quant         float
	, p_cenaEd        float
	, p_id_scet       integer
	, p_currency_rate float
)
returns integer
begin
	declare v_updated integer;


	set v_updated = update_count_remote(p_server_new, 'scet', 'summa_sale'
		, convert(varchar(20), p_currency_rate * p_quant * p_cenaEd)
		, 'id = ' + convert(varchar(20), p_id_scet)
	);
	set v_updated = update_count_remote(p_server_new, 'scet', 'summa_salev'
		, convert(varchar(20), p_quant * p_cenaEd)
		, 'id = ' + convert(varchar(20), p_id_scet)
	);
	set wf_scet_price_changed = v_updated;

end;
