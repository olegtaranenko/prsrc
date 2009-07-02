if exists (select 1 from sysprocedure where proc_name = 'wi_reset_issue_marker') then
	drop function wi_reset_issue_marker;
end if;


CREATE function wi_reset_issue_marker (
) returns integer
begin
	set @issueMarker = null;
	set wi_reset_issue_marker = 1;
end;
