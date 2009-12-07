if exists (select '*' from sysprocedure where proc_name like 'wf_next_numorder') then  
	drop procedure wf_next_numorder;
end if;

create 
	function wf_next_numorder() returns integer
begin
	declare sys_numorder integer;
	declare sys_numorder_c varchar(10);
	declare sys_year_i integer;
	declare sys_mmdd char(4);
	declare sys_number_c varchar(4);
--	declare sys_number_i integer;

	declare now_year_ln integer;
	declare now_date char(6);
	declare now_year_i integer;
	declare now_year char(2);
	declare now_mmdd char(4);
	declare v_new_base integer;


	-- по умолчанию в том же дне
	set v_new_base = 0;

	-- locking to prevent the concurrent modification
	update system set lastPrivatNum = lastPrivatNum;

	select lastPrivatNum into sys_numorder from system;
	set sys_numorder_c = convert(varchar(10), sys_numorder);

	set now_date = convert(char(6), now(), 12); -- 050716 yymmdd
	set now_year = substring(now_date, 1, 2);
	set now_year_i = convert(integer, now_year); --5 или 10 если 2010-й год
	set now_year_ln = char_length(convert(char(2), now_year_i)); --1 или 2

	-- —тандарна€ маска номера YMMDDnn[n..] 
	 
	set sys_year_i = convert(integer, substring(sys_numorder_c, 1, now_year_ln));
	if (sys_year_i != now_year_i) then
		-- ѕереход на новый год
		set v_new_base = 1;
		-- ”честь переход с 31.12.2009 на 01.01.2010
		-- измен€етс€ длина шаблона номера счета
		--if sys_year_i = 9 and now_year = 10 then
			--??? set v_year_now = 0;
		--end if;
	end if;

	
	set sys_mmdd = substring (sys_numorder, now_year_ln + 1, 4);
	set now_mmdd = substring (now_date, 3, 4);
	if sys_mmdd != now_mmdd then
		set v_new_base = 1;
	end if;

	if v_new_base = 0 then
		set sys_number_c = substring (sys_numorder_c, now_year_ln + 5);
		set sys_number_c = convert(varchar(3), convert(integer, sys_number_c) + 1);
		if char_length(sys_number_c) = 1 then
			set sys_number_c = '0' + sys_number_c;
		end if;
		set wf_next_numorder = convert(char(2),sys_year_i) + sys_mmdd + sys_number_c;
	else 
		set wf_next_numorder = convert(char(2),now_year_i) + now_mmdd + '01';
	end if;

	update system set lastPrivatNum = wf_next_numorder;

end;

