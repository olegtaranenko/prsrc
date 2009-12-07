if exists (select '*' from sysprocedure where proc_name like 'wf_next_numdoc') then  
	drop procedure wf_next_numdoc;
end if;

create 
	function wf_next_numdoc() returns integer
begin
	declare sys_numdoc integer;
	declare sys_numdoc_c varchar(10);
	declare sys_year_i integer;
	declare sys_mmdd char(4);
	declare sys_number_c varchar(4);
--	declare sys_number_i integer;

	declare now_year_ln integer;
	declare now_date char(6);
	declare now_year_i integer;
	declare now_year char(2);
	declare now_mmdd char(4);
	declare now_m char(1);
	declare v_new_base integer;


	-- по умолчанию в том же дне
	set v_new_base = 0;

	-- locking to prevent the concurrent modification
	update system set lastDocNum = lastDocNum;
	select lastDocNum into sys_numdoc from system;
	set sys_numdoc_c = convert(varchar(10), sys_numdoc);

	set now_date = convert(char(6), now(), 12); -- 050716 yymmdd
	set now_year = substring(now_date, 1, 2);
	set now_year_i = convert(integer, now_year); --5 или 10 если 2010-й год
	set now_year_ln = char_length(convert(char(2), now_year_i)); --1 или 2

	-- —тандарна€ маска номера YMMDDnn[n..] 
	 
	set sys_year_i = convert(integer, substring(sys_numdoc_c, 1, now_year_ln));
	if (sys_year_i != now_year_i) then
		-- ѕереход на новый год
		set v_new_base = 1;
		-- ”честь переход с 31.12.2009 на 01.01.2010
		-- измен€етс€ длина шаблона номера счета
		--if sys_year_i = 9 and now_year = 10 then
			--??? set v_year_now = 0;
		--end if;
	end if;

	
	set sys_mmdd = substring (sys_numdoc, now_year_ln + 1, 4);
	set now_m = convert(char(1), 2+convert(integer, convert(char(1), substring(now_date,3,1))));
	set now_mmdd = now_m + substring(now_date, 4, 3);
	if sys_mmdd != now_mmdd then
		set v_new_base = 1;
	end if;

	if v_new_base = 0 then
		set sys_number_c = substring (sys_numdoc_c, now_year_ln + 5);
		set sys_number_c = convert(varchar(3), convert(integer, sys_number_c) + 1);
		if char_length(sys_number_c) = 1 then
			set sys_number_c = '0' + sys_number_c;
		end if;
		set wf_next_numdoc = convert(char(2),sys_year_i) + sys_mmdd + sys_number_c;
	else 
		set wf_next_numdoc = convert(char(2),now_year_i) + now_mmdd + '01';
	end if;

	update system set lastDocNum = wf_next_numdoc;


end;


