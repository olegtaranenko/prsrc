if exists (select 1 from sysevent where event_name = 'aReport') then
	drop event aReport;
end if;

create event aReport
schedule 
    start time '23:55' 
    every 24 hours on ('Monday','Tuesday','Wednesday','Thursday','Friday')
handler
begin
		call wf_areprot_calculate(now(), 1, 1);
end;


