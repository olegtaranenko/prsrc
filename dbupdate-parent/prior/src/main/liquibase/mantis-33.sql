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

	set c_status_close_id = 6;  -- ������

	select sysname, invCode into remoteServerOld, v_invcode from GuideVenture where ventureId = old_name.ventureId;
	set v_currency_rate = new_name.rate;
	select id_cur into v_id_cur from system;

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
				set new_name.invoice = '���� ?';
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

			-- ����������� ��������� ���������, ��������� � �������
			--select total_account into v_total_account_date from system;

			-- ��� ����� ������ ������ ��� ��� �������, ������� ����� �������� �� ����� ������� ����� �� ������������
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

			
			-- �����, ������� ����� ������ � ���.����� ����������
			-- �.�. ��� �������� ���, ��� ������ �����

			-- ��������� ����� ��� ���������
			-- ������� ���������, ��� �� ������������� ��� ���
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
				-- ����� � ����������
				return;
			end if;
	    
			-- ���� ����� ��� ��������� "������"
			select id_inv into v_id_inv from sGuideNomenk where nomNom = '���';

			-- ������� ������� �� ����, ��� ����� ������ ��� ����.
			-- ��� ����� ��������� ��� ��������� ��������� ������.

			if abs(new_name.ordered) < 0.001 then
				call delete_remote(remoteServerOld, 'scet'
					, 'id_jmat = ' + convert(varchar(20), v_id_jscet) + ' and id_inv = ' + convert(varchar(20), v_id_inv)
				);
				return;
			end if;

			set v_id_scet = null;
			set v_id_scet = select_remote(remoteServerOld, 'scet', 'id', 'id_jmat = ' + convert(varchar(20), v_id_jscet) + ' and id_inv = ' + convert(varchar(20), v_id_inv));
			set v_cenaEd = 1; -- ���� ������ - 1 ��
			if v_id_scet is not null then

				-- ������ ����� ������
				set v_updated = wf_scet_price_changed(remoteServerOld, new_name.ordered, v_cenaEd, v_id_scet, v_currency_rate);

			else
				-- ������ ��� ����� ��� ���� => ����� ��������
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
		-- �������� ������������
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
-- �������� ����(��� � ��������) � ��� ���� ������� ��� ��������� ����� � ������.
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


if exists (select 1 from systriggers where trigname = 'wf_update_nomenk' and tname = 'sDmcRez') then 
	drop trigger sDmcRez.wf_update_nomenk;
end if;

create TRIGGER wf_update_nomenk before update on
sDmcRez
referencing old as old_name new as new_name
for each row
begin
	declare v_id_scet integer;
	declare remoteServerNew varchar(32);

	declare v_cenaEd float;
	declare v_quantity float;
	declare v_perList float;
	declare v_currency_rate float;
	
	set v_id_scet = old_name.id_scet;
	  
	select v.sysname
		, n.perList 
	into remoteServerNew
		, v_perList 
	from BayOrders o
	join GuideVenture v on o.ventureId = v.ventureId and v.standalone = 0
	join sGuideNomenk n on n.nomNom = old_name.nomNom
	where numOrder = old_name.numDoc;


	if remoteServerNew is not null then
		if update(quantity) or update(intQuant) then
			set v_currency_rate = system_currency_rate();
			set v_quantity = round(new_name.quantity/v_perList, 2);
			call update_remote(remoteServerNew, 'scet', 'summa_sale'
				, convert(varchar(20), v_currency_rate * v_quantity * new_name.intQuant)
				, 'id = ' + convert(varchar(20), v_id_scet)
			);
			call update_remote(remoteServerNew, 'scet', 'summa_salev'
				, convert(varchar(20), v_quantity*new_name.intQuant)
				, 'id = ' + convert(varchar(20), v_id_scet)
			);
        end if;
		if update(quantity) then
			call update_remote(remoteServerNew, 'scet', 'kol1', convert(varchar(20), v_quantity), 'id = ' + convert(varchar(20), v_id_scet));
		end if;
	end if;

end;
	



if exists (select '*' from sysprocedure where proc_name like 'wf_insert_scet') then  
	drop procedure wf_insert_scet;
end if;

create function wf_insert_scet (
	  p_servername varchar(20)
	, p_id_jscet integer
	, p_id_inv integer
	, p_quant float
	, p_cena float
	, p_date date
	, in p_rate float
)
returns integer
begin
	declare v_id_scet integer;
	declare v_fields varchar(255);
	declare v_values varchar(2000);
	declare scet_nu integer;
--	declare v_currency_rate float;
	declare v_datev varchar(20);
	declare v_id_cur integer;


--	set p_quant = round(p_quant, 2);
--	set p_cena = round(p_cena, 2);

 
  if p_servername is not null and p_id_jscet is not null then
//	execute immediate 'select max(nu)+1 into scet_nu from scet_' + p_servername + ' where id_jmat = ' + convert(varchar(20), p_id_jscet);

	-- �������� ��������� ���������� ����� ����� ���.����
	set scet_nu = select_remote(
		p_servername
		, 'scet'
		, 'max(nu)+1'
		, 'id_jmat = ' + convert(varchar(20), p_id_jscet)
	);

	set scet_nu = isnull(scet_nu, 1);

	-- �� ������ �����, ��������, ��� � ����������� ������ �����, � � ������ - ��
	set v_id_cur = system_currency();

--	execute immediate 'call slave_currency_rate_' + p_servername + '(v_datev, v_currency_rate, p_date, v_id_cur )';
	
	set v_fields = '
		 id_jmat
		,id_inv
		,kol1
		,nu
		,summa_sale
		,summa_salev
	';

	set v_values = 
		convert(varchar(20), p_id_jscet)
		+', '+ convert(varchar(20), p_id_inv)
		+', '+ convert(varchar(20), p_quant)
		+', '+ convert(varchar(20), scet_nu)
		+', '+ convert(varchar(20), round(p_quant*p_cena * p_rate, 2))
		+', '+ convert(varchar(20), round(p_quant*p_cena, 2))
	;
	--message 'p_cena = ', p_cena to client;
	--message 'p_quant = ', p_quant to client;
	--message 'v_values = ', v_values to client;

	-- ��������� � ������������� ���� ������
	set v_id_scet = insert_count_remote(p_servername, 'scet', v_fields, v_values);

	return v_id_scet;
  end if;
  return null;

end;




if exists (select '*' from sysprocedure where proc_name like 'wf_set_invoice_detail') then  
	drop procedure wf_set_invoice_detail;
end if;


create procedure wf_set_invoice_detail (
			p_servername varchar(20)
			, p_id_jscet integer
			, p_numOrder integer
			, p_date date
			, p_rate float
)
begin
-- ��������� �������������� �������� ������ ������
-- � ���������� ����� � ������������� ���� �������
-- ��� ����� �������, ���� � ����� ������� 
-- ������� ��������, � ������ ����� ��������� �����������,
-- ����� ������� ���� ����� ������ ������.

	declare v_id_scet integer;
	declare v_id_inv integer;
	declare is_variant integer;
	declare v_id_variant integer;
	declare is_uslug integer;
	declare v_quant float;
	declare v_perList float;

	set is_uslug = 1; // ������������ ����������, ��� ��


	for c_nomenk as n dynamic scroll cursor for
		select 
			  p.nomNom as r_nomNom
			, p.quant as r_quant
			, p.cenaEd as r_cenaEd
		from xPredmetybynomenk p
		where p.numOrder = p_numOrder
	do
	    set is_uslug = 0; -- ���� �������� � ������, ������ �� ������

		select id_inv, perList into v_id_inv, v_perList from sGuideNomenk where nomnom = r_nomNom;
		
		set v_id_scet = 
			wf_insert_scet(
				p_servername
				, p_id_jscet
				, v_id_inv
				, r_quant / v_perList
				, r_cenaEd * v_perList
				, p_date
				, p_rate
			);
		update xPredmetyByNomenk set id_scet = v_id_scet where current of n;

	end for;


	for c_izd as i dynamic scroll cursor for
		select 
			  prId as r_prId
			, prExt as r_prExt
			, quant as r_quant
			, cenaEd as r_cenaEd
		from xPredmetyByIzdelia p
		where p.numOrder = p_numOrder
	do

	    set is_uslug = 0; -- ���� �������� � ������, ������ �� ������
		select id_inv into v_id_inv from sGuideProducts where prId = r_prId;

		-- �������, �������� �� ������� ����������?
		
		select count(*) into is_variant from sVariantPower where productId = r_prId;
		if is_variant = 1 then
			-- ���� �/��� ��������� ������� � Inv
			set v_id_variant = wf_get_variant_id(p_numOrder, r_prId, r_prExt);
			select id_inv into v_id_inv 
			from sGuideComplect 
			where 
				id_variant = v_id_variant;
		end if;

		set v_id_scet = 
			wf_insert_scet(
				p_servername
				, p_id_jscet
				, v_id_inv
				, r_quant
				, r_cenaEd
				, p_date
				, p_rate
			);

		update xPredmetyByIzdelia set id_scet = v_id_scet, id_inv = v_id_inv where current of i;
	end for;  -- ����� �� ��������

	select ordered into v_quant from orders where numorder = p_numOrder;
	if is_uslug = 1 and abs(v_quant) > 0.001 then
		-- ���� ����� ��� ��������� "������"
		select id_inv into v_id_inv from sGuideNomenk where nomNom = '���';


		set v_id_scet = 
			wf_insert_scet(
				p_servername
				, p_id_jscet
				, v_id_inv
				, 1 // quant
				, v_quant//r_cenaEd
				, now()//p_date
				, p_rate
			);

	end if;


end;


if exists (select 1 from systriggers where trigname = 'wf_insert_izd' and tname = 'xPredmetyByIzdelia') then 
	drop trigger xPredmetyByIzdelia.wf_insert_izd;
end if;

create TRIGGER wf_insert_izd before insert on
xPredmetyByIzdelia
referencing new as new_name
for each row
begin
	declare v_id_scet integer;
	declare v_id_jscet integer;
	declare v_id_inv integer;
--	declare v_ventureid integer;
	declare remoteServerNew varchar(32);
	declare v_invCode varchar(10);
	declare v_fields varchar(255);
	declare v_values varchar(2000);
	declare v_date date;
	declare v_rate float;
 
	select id_jscet, inDate, sysname, invCode, o.rate
	into v_id_jscet, v_date, remoteServerNew, v_invcode, v_rate
		from orders o
		join GuideVenture v on o.ventureId = v.ventureId and v.standalone = 0
		where numOrder = new_name.numOrder;

	select id_inv into v_id_inv 
		from sGuideProducts where prId = new_name.prId;
  
	if remoteServerNew is not null and v_id_jscet is not null then
		set v_id_scet =	
			wf_insert_scet (
				remoteServerNew
				, v_id_jscet
				, v_id_inv
				, new_name.quant
				, new_name.cenaEd
				, v_date
				, v_rate
			);
		set new_name.id_scet = v_id_scet;
		set new_name.id_inv = v_id_inv;
	end if;
end;




if exists (select 1 from systriggers where trigname = 'wf_insert_nomenk' and tname = 'sDmcRez') then 
	drop trigger sDmcRez.wf_insert_nomenk;
end if;

create TRIGGER wf_insert_nomenk before insert on
sDmcRez
referencing new as new_name
for each row
begin
	declare v_id_scet integer;
	declare v_id_jscet integer;
	declare v_id_inv integer;
	declare remoteServerNew varchar(32);
	declare v_invCode varchar(10);
	declare v_date date;
	declare v_cenaEd float;
	declare v_quantity float;
	declare v_perList float;
	declare v_rate float;


--	message 'sDmcRez.wf_insert_nomenk' to client;
	select 
		o.id_jscet, o.inDate  
		, v.sysname, v.invCode
		, n.id_inv, n.perList 
		, o.rate
	into 
		v_id_jscet, v_date 
		, remoteServerNew, v_invcode
		, v_id_inv, v_perList 
		, v_rate
	from BayOrders o
	left join GuideVenture v on v.ventureid = o.ventureid and v.standalone = 0
	join sGuideNomenk n on n.nomNom = new_name.nomNom
	where o.numOrder = new_name.numDoc;


	set v_cenaEd = new_name.intQuant;
	set v_quantity = new_name.quantity / v_perList;

--	select id_inv into v_id_inv from sGuideNomenk where nomNom = new_name.nomNom;

--	select sysname, invCode into remoteServerNew, v_invcode from GuideVenture where ventureId = v_ventureId;

	if remoteServerNew is not null and v_id_jscet is not null then
	  -- �����, ������� ����� ������ � ���.����� ����������
	  -- �.�. ��� �������� ���, ��� ������ �����
		set new_name.id_scet = 
			wf_insert_scet(
				remoteServerNew
				, v_id_jscet
				, v_id_inv
				, v_quantity
				, v_cenaEd
				, v_date
				, v_rate
			);
	end if;
	  
end;



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

	set c_status_close_id = 6;  -- ������
	select sysname into remoteServerOld from GuideVenture where ventureId = old_name.ventureId;

	if update(ventureId) then
		if new_name.ventureId = 0 then
			set new_name.ventureid = null;
		end if;
		if isnull(old_name.ventureId, 0) != isnull(new_name.ventureId, 0) then
			if remoteServerOld is not null then
				call delete_remote(remoteServerOld, 'jscet', 'id = ' + convert(varchar(20), old_name.id_jscet));
				call delete_remote(remoteServerOld, 'scet', 'id_jmat = ' + convert(varchar(20), old_name.id_jscet));
				set new_name.invoice = '���� ?';
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
		-- �������� ������������
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





if exists (select '*' from sysprocedure where proc_name like 'wf_set_bay_detail') then  
	drop procedure wf_set_bay_detail;
end if;

create procedure wf_set_bay_detail (
			p_servername varchar(20)
			, p_id_jscet integer
			, p_numOrder integer
			, p_date date
			, in p_rate float
)
begin
-- ��������� �������������� �������� bay-������ ������
-- � ���������� ����� � ������������� ���� �������
-- ��� ����� �������, ���� � ����� ������� 
-- ������� ��������, � ������ ����� ��������� �����������,
-- ����� ������� ���� ����� ������ ������.

	declare v_id_scet integer;
	declare v_id_inv integer;
	declare is_variant integer;
	declare v_id_variant integer;
	declare v_quant float;

	for c_nomenk as nn dynamic scroll cursor for
		select 
			  p.nomNom as r_nomNom
			, p.quantity as r_quantity
			, intQuant as r_cenaEd
		from sDmcRez p
		where p.numDoc = p_numOrder
	do

		select 
			n.id_inv
			, r_quantity / n.perList
		into 
			v_id_inv
			, v_quant
		from 
			sGuideNomenk n
		where
			n.nomNom = r_nomNom;


		set v_id_scet = 
			wf_insert_scet(
				p_servername
				, p_id_jscet
				, v_id_inv
				, v_quant
				, r_cenaEd
				, p_date
				, p_rate
			);
		update sDmcRez set id_scet = v_id_scet where current of nn;

	end for;

end;


if exists (select 1 from systriggers where trigname = 'wf_insert_nomenk' and tname = 'xPredmetyByNomenk') then 
	drop trigger xPredmetyByNomenk.wf_insert_nomenk;
end if;

create TRIGGER wf_insert_nomenk before insert on
xPredmetyByNomenk
referencing new as new_name
for each row
begin
	declare v_id_scet integer;
	declare v_id_jscet integer;
	declare v_id_inv integer;
	declare v_ventureid integer;
	declare remoteServerNew varchar(32);
	declare v_invCode varchar(10);
	declare v_fields varchar(255);
	declare v_values varchar(2000);
	declare scet_nu integer;
	declare v_date date;
	declare v_perList float;
	declare v_rate float;

	select id_jscet, ventureId, inDate, rate
	into v_id_jscet, v_ventureId, v_date, v_rate
	from orders 
	where numOrder = new_name.numOrder;
	select id_inv, perList into v_id_inv, v_perList from sGuideNomenk where nomNom = new_name.nomNom;
	select sysname, invCode into remoteServerNew, v_invcode from GuideVenture where ventureId = v_ventureId;

	if remoteServerNew is not null and v_id_jscet is not null then
	  -- �����, ������� ����� ������ � ���.����� ����������
	  -- �.�. ��� �������� ���, ��� ������ �����
		set new_name.id_scet = 
			wf_insert_scet (
				remoteServerNew
				, v_id_jscet
				, v_id_inv
				, new_name.quant / v_perList
				, new_name.cenaEd
				, v_date
				, v_rate
			);
	end if;
	  
end;



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
//		+ ', id_kad1'
//		+ ', id_kad_bux'
//		+ ', id_s_bank'
		;

	--message 'v_fields  = ', v_fields to client;
	
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

	--message 'v_values  = ', v_values to client;

	if p_id_dest is not null then
		set v_fields = v_fields
			+ ', id_d'
			+ ', id_d_cargo'
		;
		set v_values = v_values	
			+ ', ' + convert(varchar(20), p_id_dest)
			+ ', ' + convert(varchar(20), p_id_dest)
		;
	end if;

	set r_id = insert_count_remote(remoteServerNew, 'jscet', v_fields, v_values);
end;



if exists (select '*' from sysprocedure where proc_name like 'wf_jscet_handle') then 
	drop function wf_jscet_handle;
end if;

// id �������������� ����� ��� ������
create function wf_jscet_handle (
	// �����, ������� ������ ���� ������� � ��������� ����
	  p_numorder integer			
	, in p_id_jscet_new integer default null
) returns integer
begin
	// ��������� ������ ������� ����� ���� �������� �� ���
	declare old_invoice varchar(10);
	declare old_ventureId integer;
	declare old_id_jscet integer;
	declare old_invCode varchar(20);
	declare old_server varchar(20);
	declare v_nu_jscet varchar(50);
	declare v_id_jscet integer;
	declare v_id_dest integer;
	declare v_rate float;

	--message 'p_numorder = ', p_numorder to client;
	--message 'p_id_jscet_new = ', p_id_jscet_new to client;

	select invoice, id_jscet, o.ventureId, v.invCode, v.sysname, f.id_voc_names, o.rate
	into old_invoice, old_id_jscet, old_ventureId, old_invCode, old_server, v_id_dest, v_rate
	from orders o
		join guideventure v on v.ventureId = o.ventureId
		join guidefirms f on f.firmid = o.firmid
	where numorder = p_numorder;

	if old_ventureId is null then
		return;
	end if;

	if p_id_jscet_new is not null then
		// ������ ����������� ���
		set v_id_jscet = p_id_jscet_new;
		set v_nu_jscet = select_remote (old_server, 'jscet', 'nu', 'id = ' + convert(varchar(20), p_id_jscet_new));
//		set out_invoice = old_invCode + convert(varchar(20), v_nu_jscet);
	else
		// ��������� ������ � ��������� ����
		call put_jscet (v_id_jscet, v_nu_jscet, old_server, p_numOrder, v_id_dest, old_invoice, v_rate);
	end if;

	update orders set id_jscet = v_id_jscet where numOrder = p_numorder;
	update orders set invoice = old_invCode + convert(varchar(20), v_nu_jscet) where numOrder = p_numorder;

	// ����� �������� ������ �� ������ �����, ������� ��������� 
	// � ������ � ��������� �� � ����� ����
	--message ' old_server = ', old_server to client;
	--message ' v_id_jscet = ', v_id_jscet to client;
	--message ' p_numOrder = ', p_numOrder to client;
	call wf_move_invoice_detail (old_server, v_id_jscet, p_numOrder);

	// ��������� ���������� 
	// ������ ������� ��� ������ � ������� �����
	call call_remote(old_server, 'slave_renu_scet', v_id_jscet);
	call call_remote(old_server, 'slave_renu_scet', old_id_jscet);

//	return convert(integer, v_nu_jscet);
	return v_id_jscet;
end;


