if exists (select 1 from sysprocedure where proc_name = 'n_get_analysid') then
	drop function n_get_analysid;
end if;


CREATE function n_get_analysid (
	  p_rowid    integer
	  , p_columnid integer 
) returns integer
begin
	select id into n_get_analysid from nAnalys where byrow = p_rowid and bycolumn = p_columnid;
end;
