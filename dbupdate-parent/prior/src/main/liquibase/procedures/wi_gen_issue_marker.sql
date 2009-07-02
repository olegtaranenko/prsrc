if exists (select 1 from sysprocedure where proc_name = 'wi_gen_issue_marker') then
	drop function wi_gen_issue_marker;
end if;


CREATE function wi_gen_issue_marker (
	  p_prefix varchar(5) default ''
) returns varchar(32)
begin
	set wi_gen_issue_marker = p_prefix + '|' + convert(varchar(40), now(), 109);

	set @issueMarker = wi_gen_issue_marker;
end;
