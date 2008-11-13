if exists (select 1 from sysprocedure where proc_name = 'wf_territory_catalog') then
	drop procedure wf_territory_catalog;
end if;

CREATE procedure wf_territory_catalog (
	p_firmAlso integer default 0
)
begin
	declare v_ord_table varchar(64);
	declare p_table_name varchar(64);    
	declare p_id_name varchar(64);       
	declare p_parent_id_name varchar(64);
	declare p_order_by_name varchar(256);

	set p_table_name = 'bayRegion';
	set p_id_name = 'regionId';
	set p_parent_id_name = 'territoryId';
	set p_order_by_name = 'region';

	set v_ord_table = get_tmp_ord_table_name(p_table_name);
	execute immediate 'create table ' + v_ord_table + ' (id integer, ord integer)';

	call wf_sort_klassificator(p_table_name, p_id_name, p_parent_id_name, p_order_by_name);

--	select * from #bayRegion_ord order by 1;

	if p_firmAlso = 1 then
		-- регионы вместе с фирмами.
		select r.regionId, r.region, r.territoryId as territoryId, o.ord, f.firmId, f.name as firmName
		from bayRegion r 
		join #bayRegion_ord o on o.id = r.regionId
		left join bayGuideFirms f on f.regionId = r.regionId
		where isnull(region, '') != ''
		order by o.ord, r.region, firmName, f.firmId;
	else
		-- только регионы, как было изначально.
		select r.regionId, r.region, r.territoryId as territoryId, o.ord
		from bayRegion r 
		join #bayRegion_ord o on o.id = r.regionId
		where isnull(region, '') != ''
		order by o.ord, r.region;
	end if;

	execute immediate get_tmp_ord_drop_sql(v_ord_table);
end;



