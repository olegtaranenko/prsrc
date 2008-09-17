begin
	declare tsql varchar(256);

	for c_events as cur dynamic scroll cursor for
		select * from sysevent
	do
		message event_name to client;
		set tsql = 'alter event "' + event_name + '" disable';
		execute immediate tsql;
	end for
end;
