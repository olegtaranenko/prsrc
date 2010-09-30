begin
	create table #tmp (numorder int, outdatetime datetime);

	insert #tmp (numorder, outdatetime)
	select numorder, max(outdatetime)
	from ordersequip
	group by numorder;

	UPDATE orders join #tmp t on t.numorder = orders.numorder
	set orders.Outdatetime = t.Outdatetime
	where orders.werkid = 2;
end;
