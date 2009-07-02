if exists (select 1 from sysprocedure where proc_name = 'wi_check_issue_action') then
	drop procedure wi_check_issue_action;
end if;


CREATE procedure wi_check_issue_action (
	  p_issueId integer
) 
begin
	select it.description, it.action, idt.sysname as issueDetail, d.value as detailValue
	from iIssueType it
	join iBusinessIssue bi on bi.issueTypeId = it.issueTypeId and bi.issueId = p_issueId
	left join iIssueDetail d on d.issueId = bi.issueId
	left join iIssueDetailType idt on idt.issueTypeId = it.issueTypeId and idt.issueDetailTypeId = d.issueDetailTypeId
	where idt.issueClass = 'showMsgBox' 
	;
end;
