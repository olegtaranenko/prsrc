ALTER FUNCTION "DBA"."n_check_filter" (
	  p_filterid    integer
	, p_managId     varchar(16)
) returns varchar(254)
begin
	declare v_byrow_id       integer;
	declare v_bycolumn_id    integer;
	declare v_byrow          varchar(31);
	declare v_bycolumn       varchar(31);
	declare v_passed         integer;


	set n_check_filter = 'ok';
	select byrow, bycolumn, r.name, c.name
	into v_byrow_id, v_bycolumn_id, v_byrow, v_bycolumn
	from 
		nAnalys a
	join nAnalysCategory r on r.id = a.byrow
	join nAnalysCategory c on c.id = a.bycolumn
	join nAnalysTemplate t on a.templateId = t.id
	join nFilter f on f.id = p_filterId and f.byrowid = a.byrow and f.bycolumnid = a.bycolumn
	;

	if v_byrow_id is null then
		set n_check_filter = 'Эта функция еще не реализована';
		return;
	end if;


	if v_byrow = 'firm' and v_bycolumn = 'klasses' then
		set v_passed = 0;
		for x as xc dynamic scroll cursor for
			call n_filter_params(p_filterid)
		do
			if r_itemType = 'materials' and r_isActive = 1 then
				set v_passed = 1;
			end if;
		end for;

		if v_passed != 1 then
			set n_check_filter = 'Необходимо определить группы материалов';
			return;
		end if
	end if;
end