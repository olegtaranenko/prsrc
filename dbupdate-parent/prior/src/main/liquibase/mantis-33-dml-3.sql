begin
	declare v_id integer;
	set v_id = 1;
	for sx as x dynamic scroll cursor for
		select id from ybook order by xdate for update
	do
		update ybook set id = v_id where current of x;
		set v_id = v_id + 1;
	end for;
end;
