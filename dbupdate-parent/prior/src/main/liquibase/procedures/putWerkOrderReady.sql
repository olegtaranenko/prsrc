if exists (select 1 from sysprocedure where proc_name = 'putWerkOrderReady') then
	drop procedure putWerkOrderReady;
end if;


CREATE procedure putWerkOrderReady (
	  p_numorder  integer
	, p_xDate     varchar(10)
	, p_obrazec   varchar(1)
	, p_virabotka float
	)
begin
	
	declare v_totalByEquip float;
	declare v_worktime float;
	declare v_worktimeMO float;
	declare v_koef float;

	set v_totalByEquip = 0;

	select sum(isnull(oe.worktime, 0)), sum(isnull(oe.worktimeMO,0))
		into v_worktime, v_worktimeMO
	from OrdersEquip oe
	where oe.numorder = p_numorder;

	for o as oc dynamic scroll cursor for
		select 
			  oe.equipId  as r_equipId
			, r.equipId   as r_equipIdRes
			, r.virabotka as r_virabotka
			, oe.worktime as r_worktime
			, oe.worktimeMO as r_worktimeMO
		from OrdersEquip oe 
		left join Itogi r on 
				r.equipId = oe.equipId 
			and r.numorder = oe.numorder
			and r.xdate = p_xDate 
			and r.obrazec = p_obrazec
		where oe.numorder = p_numorder
	do

		if isnull(p_obrazec, '') <> '' then
			set v_koef = r_worktimeMO / v_worktimeMO;
		else
			set v_koef = r_worktime / v_worktime;
		end if;

		if r_equipIdRes is null then
			insert Itogi (equipId, xDate, obrazec, virabotka, numorder) 
			select r_equipId, p_xDate, p_obrazec, p_virabotka * v_koef, p_numorder;
		else
			update Itogi set virabotka = virabotka + p_virabotka * v_koef
			where equipId = r_equipId and xdate = p_xDate and obrazec = p_obrazec and numorder = p_numorder;
		end if;

	end for;

end;

