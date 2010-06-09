if exists (select 1 from sysprocedure where proc_name = 'putWerkOrderReady') then
	drop procedure putWerkOrderReady;
end if;


CREATE procedure putWerkOrderReady (
	  p_numorder  integer
	, p_xDate     varchar(10)
	, p_obrazec   varchar(1)
	, p_virabotka float
	, p_equipId   integer
	, p_nevip     float
	)
begin
	
	declare v_equipId integer;

	set v_equipId = p_equipId;

	update Itogi set virabotka = virabotka + p_virabotka
	where equipId = v_equipId and xdate = p_xDate and obrazec = p_obrazec and numorder = p_numorder;

	if @@rowcount = 0 then
		insert Itogi (equipId, xDate, obrazec, virabotka, numorder) 
		select v_equipId, p_xDate, p_obrazec, p_virabotka, p_numorder;
	end if;

	if p_obrazec = '' then
		update OrdersEquip set nevip = p_nevip where numorder = p_numorder and equipId = v_equipId;	
	end if; 
end;


