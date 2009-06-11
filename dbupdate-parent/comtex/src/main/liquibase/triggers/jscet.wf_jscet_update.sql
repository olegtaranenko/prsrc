if exists (select 1 from systriggers where trigname = 'wf_jscet_update' and tname = 'jscet') then 
	drop trigger jscet.wf_jscet_update;
end if;

create TRIGGER wf_jscet_update before update order 211 on
/*
	Выставляет плательщика в базе Prior
	передает договор от одного клиента к другому, если у заказа(счета) меняется дебитор.
	Заботится о том, чтобы один и тот же договор не был присвоен больше, чем одному счету.
	Если же вдруг пользователь пытается это сделать, то просто это поле обнуляется.
*/
jscet
referencing old as old_name new as new_name
for each row
begin
	declare no_echo integer;
	declare v_is_orders varchar(10);
	declare v_id_jdog integer;
	declare v_check_uniqueness integer;
	
	set no_echo = 0;

  	begin
		select @prior_jscet into no_echo; 
	exception 
		when other then
			set no_echo = 0;
	end;

	if no_echo = 1 then
		return;
	end if;

//	message 'TRIGGER wf_jscet_update:: no_echo = ', no_echo to client;
	
	if update(id_d) then
		// передаем как эстафетную палочку договор новому клиенту
		set v_id_jdog = old_name.id_jdog;
		if v_id_jdog is not null then
			// поскольку комтех обнуляет поле jscet.id_jdog (в общем-то делает верно)
			// то нам нужно еще получить обратный ход от приора после updating id_bill.
			update jdog set id_post = new_name.id_d where jdog.id = v_id_jdog
		end if;

		// выяснить это продажи или заказ
		set v_is_orders = admin.select_remote('prior'
			,'orders'
			,'count(*)'
			,'id_jscet = ' + convert(varchar(20), old_name.id)
		);

//		message 'v_is_orders = ', v_is_orders to client;

		if v_is_orders = '0' then
			// нет такого счета в Заказах
			call admin.update_remote (
				'prior'
				, 'bayOrders'
				, 'id_bill'
				, convert(varchar(20), new_name.id_d)
				, 'id_jscet = ' + convert(varchar(20), old_name.id)
			);
		else
			call admin.update_remote (
				'prior'
				, 'orders'
				, 'id_bill'
				, convert(varchar(20), new_name.id_d)
				, 'id_jscet = ' + convert(varchar(20), old_name.id)
			);
		end if;

	end if;

	if update(id_jdog) and new_name.id_jdog > 0 then
		select count(*) into v_check_uniqueness from jscet s where s.id_jdog = new_name.id_jdog;
		if v_check_uniqueness >= 1 then
			set new_name.id_jdog = 0;
		end if;
	end if;

end;

