if exists (select 1 from sysprocedure where proc_name = 'wi_add_issue_attribute') then
	drop function wi_add_issue_attribute;
end if;


CREATE function wi_add_issue_attribute (
	  p_issueId int
	, p_key     varchar(63)
	, p_value   long varchar
	, p_seqOrder integer default 0
) returns int
begin

	declare v_seqOrder integer;
	declare v_issueDetailTypeId integer;
	declare v_issueTypeId integer;

	if p_seqOrder = 0 then
		select max(seqOrder) + 1 into v_seqOrder from iIssueDetail where issueId = p_issueId;
		if v_seqOrder is null then
			set v_seqOrder = 1;
		end if;
	else
		set v_seqOrder = p_seqOrder;
	end if;


	select idt.issueDetailTypeId, bi.issueTypeId
		into v_issueDetailTypeId, v_issueTypeId
	from iBusinessIssue bi
	left join iIssueDetailType idt on bi.issueTypeId = idt.issueTypeId and idt.sysname = p_key
	where bi.issueId = p_issueId;


	if v_issueDetailTypeId is null then
		-- lazy inserting of the Issue Detail Type
		insert into iIssueDetailType (issueTypeId, sysname) values (v_issueTypeId, p_key);
		set v_issueDetailTypeId = @@identity;
	end if;



	insert into iIssueDetail (issueId, [value], seqOrder, issueDetailTypeId)
	values(p_issueId, p_value, v_seqOrder, v_issueDetailTypeId);
	set wi_add_issue_attribute = @@identity;

end;
