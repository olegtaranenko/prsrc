if exists (select 1 from sysprocedure where proc_name = 'wi_get_msgcode') then
	drop function wi_get_msgcode;
end if;


CREATE function wi_get_msgcode (
	  p_issueId int
) returns int
begin

	select it.msgCode into wi_get_msgcode
	from iIssueType it
	join iBusinessIssue bi on bi.issueTypeId = it.issueTypeId
	where bi.issueId = p_issueId;
end;
