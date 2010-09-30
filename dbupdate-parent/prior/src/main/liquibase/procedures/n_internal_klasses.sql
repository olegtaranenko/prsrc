if exists (select 1 from sysprocedure where proc_name = 'n_internal_klasses') then
	drop procedure n_internal_klasses;
end if;


CREATE procedure n_internal_klasses (
	  p_begin         date
	, p_end           date
	, p_table_name    varchar(64)
	, p_firmId        integer
	, p_klassId       integer
)
begin
	declare v_region_flag integer;
	declare v_bayStatus_flag integer;
	declare v_tool_flag      integer;
	declare v_sql long varchar;

	declare v_ord_table varchar(64);
	declare p_id_name varchar(64);       
	declare p_parent_id_name varchar(64);
	declare p_order_by_name varchar(256);

	declare v_cnt integer;

	declare v_material_flag integer;
	declare v_firmId           integer;


	if isnull(p_klassId, 0) != 0 then
		delete from #materials where klassId != p_klassId;
		set v_material_flag = 1;
	else
		select count(*) into v_material_flag 	from #materials where isActive = 1;
	end if;

	select count(*) into v_region_flag   	from #regions   where isActive = 1;

	
	select toolId into v_tool_flag  from #tool;
	if v_tool_flag is null then
		set v_tool_flag = 0
	end if;


	
	select bayStatusId into v_bayStatus_flag  from #bayStatus;

	if v_bayStatus_flag is null then
		set v_bayStatus_flag = 0
	end if;



	set v_firmId = p_firmId;
	if v_firmId = 0 then
		set v_firmId = null;
	end if;

--	message 'v_region_flag = ', v_region_flag to client;
--	message 'v_material_flag = ', v_material_flag to client;
--	message 'v_firmId = ', v_firmId to client;
--	message 'p_klassId = ', p_klassId to client;
--	message 'p_end = ', p_end to client;

--	message 'p_begin = ', p_begin to client;
--	message 'p_end = ', p_end to client;

	insert into #sale_item (
		 numorder
		,nomnom
		,materialQty
		,sm
		,inDate
		,firmId      
		,klassid
	)
	select
		  o.numorder as numorder
		, i.nomnom
		, i.quant / n.perlist as materialQty
		, (i.quant / n.perlist) * i.cenaEd as sm
		, o.inDate
		, o.firmId
		, n.klassid
	from itemSellOrde i
	join bayorders o on o.numorder = i.numorder 
	join sguidenomenk n on i.nomnom = n.nomnom
	join bayGuideFirms f on o.firmId = f.firmId and isnull(v_firmId, f.firmId) = f.firmId
	where 
			(p_begin is null or o.indate >= p_begin) and (p_end is null or o.inDate < p_end)
		and (v_material_flag = 0 or exists (select 1 from #materials m where n.klassid = m.klassid))
		and (v_region_flag = 0 
			or exists (
				select 1 
				from #regions r 
				where r.regionid = f.regionId
			)
		)
		and (v_tool_flag = 0 
			or (v_tool_flag = -1 and not exists (
				select 1  
				from FirmTools ft
				where ft.firmId = o.firmId
				)
			)
			or (v_tool_flag > 0 and exists (
				select 1 
				from FirmTools ft
					, #tool tt
				where ft.firmId = o.firmId and tt.toolId = ft.toolId
				)
			)
		)
		and (  v_bayStatus_flag = 0 
			or (v_bayStatus_flag = -1 
				and not exists (
					select 1 
					from bayGuideFirms bf 
					where bf.firmId = o.firmId and bf.bayStatusId > 0
				)
			)
			or (v_bayStatus_flag > 0 
				and exists (
					select 1 
					from #bayStatus bs, bayGuideFirms bf 
					where o.firmId = bf.firmId and bs.bayStatusId = bf.bayStatusId
				)
			)
		)
	;
--	message 'count of #sale_item = ', @@rowcount to client;

	delete from #materials where not exists (select 1 from #sale_item i where i.klassId = #materials.klassId);

	select count(*) into v_cnt from #materials;
--	message 'count(*) from #materials = ', v_cnt to client;

	set p_id_name = 'klassId';
	set p_parent_id_name = 'parentKlassId';
	set p_order_by_name = 'klassName';
	set v_ord_table = get_tmp_ord_table_name(p_table_name);


	call wf_sort_klassificator(p_table_name, p_id_name, p_parent_id_name, p_order_by_name);

	if v_material_flag = 0 then
		insert into #periods (klassId, label)
		select k.klassId as r_klassId, k.klassName as r_klassName
		from sGuideKlass k 
		join #sGuideKlass_ord o on o.id = k.klassId
		where isnull(k.klassName, '') != ''
		order by o.ord, k.klassName;
	else
		insert into #periods (klassId, label)
		select k.klassId as r_klassId, k.klassName as r_klassName
		from sGuideKlass k 
		join #sGuideKlass_ord o on o.id = k.klassId
		join #materials m on m.klassId = k.klassId
		where isnull(k.klassName, '') != ''
		order by o.ord, k.klassName;
	end if;		



end;


