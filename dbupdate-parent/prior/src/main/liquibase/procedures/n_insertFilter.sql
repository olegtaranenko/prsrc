ALTER FUNCTION "DBA"."n_insertFilter" (
	  p_filter_name  varchar(64)
	, p_manager      char(1)
	, p_personal     integer
	, p_byrowid      integer
	, p_bycolumnid   integer
	, p_time         datetime default null
) returns integer
begin
	insert into nFilter (name, managId, personal, created, byrowid, bycolumnid)
	select 
		p_filter_name, m.managId, p_personal, isnull(p_time, now()), p_byrowid, p_bycolumnid
	from 
		GuideManag m 
	where 
		m.manag = p_manager
	;

	set n_insertFilter = @@identity;
end