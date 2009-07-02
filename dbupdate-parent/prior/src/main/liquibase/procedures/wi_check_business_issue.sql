if exists (select 1 from sysprocedure where proc_name = 'wi_check_business_issue') then
	drop function wi_check_business_issue;
end if;


CREATE function wi_check_business_issue (
	  p_issueMarker varchar(32) default ''
) returns int
begin
	select max(issueId) 
	into wi_check_business_issue 
	from iBusinessIssue
	where issueMarker = p_issueMarker;
end;
