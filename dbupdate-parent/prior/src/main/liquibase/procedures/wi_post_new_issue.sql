if exists (select 1 from sysprocedure where proc_name = 'wi_post_new_issue') then
	drop function wi_post_new_issue;
end if;


CREATE function wi_post_new_issue (
	  p_issueType   varchar(63)
	, p_application varchar(15)
	, p_msg         long varchar default null
) returns int
begin
	set wi_post_new_issue = 0;

	insert iBusinessIssue (issueTypeId, application, msg, issueMarker)
	select it.issueTypeId, p_application, p_msg, @issueMarker
	from iIssueType it
	where it.sysname = p_issueType;

	if @@rowcount > 0 then
		set wi_post_new_issue = @@identity;
	end if
end;
