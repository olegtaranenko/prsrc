ALTER FUNCTION "DBA"."n_insertItem" (
	  p_filter_id  varchar(64)
	, p_item_name    varchar(64)
	, p_active       integer
) returns integer
begin
	insert into nItem (filterId, itemTypeId, isActive) 
	select p_filter_id, it.id, p_active
	from 
		nItemType it 
	where 
		it.itemType = p_item_name
	;

	set n_insertItem = @@identity;
end