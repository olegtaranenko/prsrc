if exists (select 1 from sysprocedure where proc_name = 'n_removeItem') then
	drop procedure n_removeItem;
end if;

CREATE procedure n_removeItem(
	  p_filter_id    integer
	, p_item_name    varchar(64)
)
begin
	delete from nItem 
	from 
		nItemType it 
	where 
		it.itemType = p_item_name
		and nItem.itemTypeId = it.id
		and nItem.filterId = p_filter_id
	;

end
