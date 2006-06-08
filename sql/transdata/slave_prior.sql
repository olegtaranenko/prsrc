if exists (select 1 from sysprocedure where proc_name = 'cast_acc') then
	drop function cast_acc;
end if;

create 
	function cast_acc(
		in sc varchar(26)
		,in base integer default 2
	)
	returns varchar(26)
begin
	if sc is null then
		set sc = '';
	end if;

	set sc = trim(sc);
	return string(repeat('0', base - char_length(sc)), sc);
end;

---------------------------------------------------------------------------
---------------------------------------------------------------------------


if exists (select 1 from sysprocedure where proc_name = 'slave_put_account') then
	drop procedure slave_put_account;
end if;

create 
	procedure slave_put_account
	(
		out out_exists integer
		, inout p_sc varchar(26)
		, inout p_sub varchar(10)
		, inout p_name varchar(98)
		, inout p_desc varchar(98) 
	)
begin
/*
	declare GLOBALSETTING_CHECK_ACCOUNT integer;
	set GLOBALSETTING_CHECK_ACCOUNT  = 1;
	if GLOBALSETTING_CHECK_ACCOUNT = 0 then
		set out_exists = 1;
		return;
	end if;
*/
	
	set p_sc = cast_acc(p_sc);
	set p_sub = cast_acc(p_sub);

	select count(*) into out_exists from yGuideSchets
	where number = p_sc and subNumber = p_sub;

	if out_exists = 0 then
		insert into yGuideSchets (number, subNumber, note, subNote) values (p_sc, p_sub, p_name, p_desc);
	end if;

/*	
	set out_exists = 1;
	if out_exists = 0 and GLOBALSETTING_CHECK_ACCOUNT = 1 then
		raiserror 17001 'В базе prr не существует счета %1!/%2!'
			+ 'Проверьте согласованность планов счетов.'
		, p_sc, p_sub;
	end if;
*/
end;


if exists (select 1 from sysprocedure where proc_name = 'slave_bind_zakaz') then
	drop function slave_bind_zakaz;
end if;

create 
	procedure slave_bind_zakaz
	(
		  out v_orderNum varchar(150)
		, p_server     varchar(50) // от какого сервера
		, p_invoice    varchar(10) // номер счета, к которому нужно найти все заказы
		, in p_summa      float        // сумма в рублях
		, in p_sc_credit  varchar(10)
		, in p_id_xoz     integer default null   // можно ли делать сразу update?
	)
begin
    declare v_sep            char(1)    ;
    declare v_order_ordered  float      ;
    declare v_order_paid     float      ;
    declare v_balans_ok      integer    ;
	declare v_ordered        float      ;
    declare v_order_count    integer    ;
	declare v_invCode        varchar(10);
	declare v_ventureId      integer    ;
	declare v_summav         float      ;

    set v_balans_ok = 0;
    set v_sep = '';
    set v_orderNum = '';

    select invCode, ventureId
    into v_invCode, v_ventureId
    from guideventure 
    where sysname = p_server;

    message 'in slave_bind_zakaz_', @@servername to client;

	if v_invCode is not null and char_length(p_invoice) > 0 then
		set v_order_count = 0;
		set v_orderNum = '';
		set v_order_ordered = 0.0;
		set v_order_paid = 0.0;

		set v_summav = p_summa / system_currency_rate();


		for v_server_name as a dynamic scroll cursor for
			select numOrder
				, isnull(ordered,0) as ordered
				, isnull(paid,0) as paid 
			from orders 
			where invoice = v_invCode + p_invoice 
				and isnull(ordered, 0) != isnull(paid, 0)
			order by invoice desc
		do
			set v_order_count = v_order_count + 1;
			set v_orderNum = v_orderNum + v_sep + convert(varchar(20), numOrder);
			set v_order_ordered = v_order_ordered + ordered;
			set v_order_paid  = v_order_paid + paid;
			set v_sep = '/';
		end for;



		if v_order_count > 0 then
			if round(v_order_ordered, 2) = round(v_order_paid + v_summav, 2) then
				set v_balans_ok = 1;
			else
				set v_orderNum = 'Заказ(ы): ' + v_orderNum + '. Ошибка при контроле суммы. '
					+ ' зак-но всего='+convert(varchar(20), round(v_order_ordered, 2))
					+ ';опл-но раньше='+convert(varchar(20), round(v_order_paid, 2))
					+ ';сумма тек.оплаты='+convert(varchar(20), round(v_summav, 2))
				;
			end if;
			if p_sc_credit != '62' then
				set v_orderNum = 'Заказ(ы): ' + v_orderNum +'. Сумма совпадает, но счет Неправильный (дб. 62).';
				set v_balans_ok = 0;
			end if;
		end if;

		if v_balans_ok = 1 then
			for v_server_name as aa dynamic scroll cursor for
				select paid from orders 
				where invoice like v_invCode + p_invoice 
					and isnull(ordered, 0) != isnull(paid, 0)
				for update
			do

				UPDATE orders set paid = ordered WHERE CURRENT OF aa;

			end for;
		end if;

	-- Теперь проверяем в продажах
		if v_order_count = 0 then

			for v_server_name as b dynamic scroll cursor for
				select numOrder
					, isnull(paid,0) as paid 
				from bayorders 
				where invoice like v_invCode + p_invoice 
				order by invoice desc
			do
				-- в bayorders поле ordered не заполняется,
				-- а высчитывается динамически :-(
				select sum (d.quantity / n.perlist * d.intquant) 
				into v_ordered
				from sdmcrez d
				join sguidenomenk n on d.nomnom = n.nomnom
				where numdoc = numOrder;
				    
				if isnull(v_ordered, 0) != isnull(paid, 0) then 
					set v_order_count = v_order_count + 1;
					set v_orderNum = v_orderNum + v_sep + convert(varchar(20), numOrder);
					set v_order_ordered = v_order_ordered + v_ordered;
					set v_order_paid  = v_order_paid + paid;
					set v_sep = '/';
				end if;
			end for;
	    
	    
			if v_order_count > 0 then
				if round(v_order_ordered, 2) = round(v_order_paid + v_summav, 2) then
					set v_balans_ok = 1;
				else
					set v_orderNum = 'Заказ(ы): ' + v_orderNum + ' Ошибка при контроле суммы. '
						+ ' зак-но всего='+convert(varchar(20), round(v_order_ordered, 2))
						+ ';опл-но раньше='+convert(varchar(20), round(v_order_paid, 2))
						+ ';сумма тек.оплаты='+convert(varchar(20), round(v_summav, 2))
					;
				end if;
				if p_sc_credit != '62' then
					set v_orderNum = 'Заказ(ы): ' + v_orderNum +'. Ошибка при контроле номера счета (' + p_sc_credit + ')';
					set v_balans_ok = 0;
				end if;
			end if;
	    
			if v_balans_ok = 1 then
				for v_server_name as bu dynamic scroll cursor for
					select numorder, paid 
					from bayorders 
					where invoice = v_invCode + p_invoice 
					for update
				do
	    
					select sum (d.quantity / n.perlist * d.intquant) 
					into v_ordered
					from sdmcrez d
					join sguidenomenk n on d.nomnom = n.nomnom
					where numdoc = numOrder;
	            
					if isnull(v_ordered, 0) != isnull(paid, 0) then 
						UPDATE bayorders set paid = v_ordered WHERE CURRENT OF bu;
					end if;
	    
				end for;
			end if;

		end if;
	end if;

	if p_id_xoz is not null and char_length(v_orderNum) > 0 then
		update ybook set ordersNum = 
				if isnull(ordersNum, '') != '' 
				then ordersNum + ' ' + v_orderNum 
				else v_orderNum 
				endif
		where id_xoz = p_id_xoz and ventureId = v_ventureId;
		;
	end if;
	
end;


if exists (select 1 from sysprocedure where proc_name = 'slave_put_xoz') then
	drop procedure slave_put_xoz;
end if;

create 
	procedure slave_put_xoz
	(
		  p_server     varchar(50)
		, p_id_xoz	   integer
		, inout p_debit_sc   varchar(26)
		, inout p_debit_sub  varchar(10)
		, inout p_credit_sc  varchar(26)
		, inout p_credit_sub varchar(10)
		, p_dat        varchar(20)
		, p_sum        float
		, p_sumv       float
		, p_id_curr    integer
		, p_detail     varchar(99)
		, p_purposeId  integer
		, p_kredDebitor integer
		, p_invoice       varchar(10)
	)
begin
    declare v_ventureid integer;
    declare v_currency_rate float;
    declare v_currency float;
    declare v_date datetime;
    declare v_orderNum       varchar(150);

    if p_dat is not null and char_length(p_dat) > 0 then
	    set v_date = convert(datetime, p_dat);
	else
		set v_date = now();
	end if;

    select ventureid
    into v_ventureid
    from guideventure 
    where sysname = p_server;


	set v_currency_rate = system_currency_rate();
	set v_currency = p_sum / v_currency_rate;

	set p_debit_sc    = cast_acc (p_debit_sc   );
	set p_debit_sub   = cast_acc (p_debit_sub  );
	set p_credit_sc   = cast_acc (p_credit_sc  );
	set p_credit_sub  = cast_acc (p_credit_sub );

	call slave_bind_zakaz (v_orderNum, p_server, p_invoice, p_sum, p_credit_sc);


	insert into yBook(
		  ventureid
		, id_xoz
		, xDate
		, UEsumm
		, Debit
		, subDebit
		, Kredit
		, subKredit
		, kredDebitor
		, ordersNum
		, purposeId
		, descript
		, Note
	) values (
		  v_ventureid
		, p_id_xoz
		, v_date
		, v_currency
		, p_debit_sc
		, p_debit_sub
		, p_credit_sc
		, p_credit_sub
		, p_kredDebitor
		, v_orderNum
		, p_purposeId
		, p_detail
		, p_invoice
	);
end;


if exists (select 1 from sysprocedure where proc_name = 'slave_set_purpose') then
	drop procedure slave_set_purpose;
end if;

create 
	procedure slave_set_purpose
	(
		  p_purpose    varchar(99)
		, inout p_debit_sc   varchar(26)
		, inout p_debit_sub  varchar(10)
		, inout p_credit_sc  varchar(26)
		, inout p_credit_sub varchar(10)
		, out v_purposeId integer
	)
begin
	declare v_ventureid integer;
--	declare v_descript varchar(100);
	declare v_currentId integer;

	set p_debit_sc    = cast_acc (p_debit_sc   );
	set p_debit_sub   = cast_acc (p_debit_sub  );
	set p_credit_sc   = cast_acc (p_credit_sc  );
	set p_credit_sub  = cast_acc (p_credit_sub );


	select pId into v_purposeid from yGuidePurpose
	where 
			Debit = p_debit_sc
		and subDebit = p_debit_sub
		and Kredit = p_credit_sc
		and subKredit = p_credit_sub
		and pDescript = p_purpose;

	if v_purposeid is null then

		if not exists (select 1 from yGuidePurp where descript = p_purpose) then
			insert into yGuidePurp (descript) values (p_purpose);
		end if;
			
		-- вставляем в таблицу
		insert into yGuidePurpose (
    		Debit, subDebit, Kredit, subKredit, pDescript
		) values (
			p_debit_sc, p_debit_sub, p_credit_sc, p_credit_sub
			, p_purpose
		);

		select pId into v_purposeid from yGuidePurpose
		where 
				Debit = p_debit_sc
			and subDebit = p_debit_sub
			and Kredit = p_credit_sc
			and subKredit = p_credit_sub
			and pDescript = p_purpose;

	end if;


end;

	