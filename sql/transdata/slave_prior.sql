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
		raiserror 17001 'В базе Prior не существует счета %1!/%2!'
			+ 'Проверьте согласованность планов счетов.'
		, p_sc, p_sub;
	end if;
*/
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
	)
begin
    declare v_ventureid integer;
    declare v_currency_rate float;
    declare v_currency float;
    declare v_date datetime;
    if p_dat is not null and char_length(p_dat) > 0 then
	    set v_date = convert(datetime, p_dat);
	else
		set v_date = now();
	end if;

    select ventureid into v_ventureid from guideventure where sysname = p_server;
    if v_ventureid is null then
    	raiserror 17000, 'Сервер %1! не существует', p_server;
    end if;

	set v_currency_rate = system_currency_rate();
	set v_currency = p_sum / v_currency_rate;

	set p_debit_sc    = cast_acc (p_debit_sc   );
	set p_debit_sub   = cast_acc (p_debit_sub  );
	set p_credit_sc   = cast_acc (p_credit_sc  );
	set p_credit_sub  = cast_acc (p_credit_sub );


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
		, 0
		, ''
		, p_purposeId
		, p_detail
		, 'из Komtex'
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

	