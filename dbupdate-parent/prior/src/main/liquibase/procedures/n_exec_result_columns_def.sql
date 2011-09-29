ALTER PROCEDURE "DBA"."n_exec_result_columns_def" (
	  p_headType    integer
	, p_managId     varchar(16)
	, p_filterId    integer default null
	, p_byrow       integer default null
	, p_bycolumn    integer default null
) 
begin
	declare v_templateId    integer;
	declare v_headerId      integer;


	if isnull(p_filterId, 0) != 0 then
		select a.templateId, t.headerId
		into v_templateId, v_headerId
		from nAnalys a
		join nFilter f on f.id = p_filterId and a.byrow = f.byrowId and a.bycolumn = f.bycolumnid
		join nAnalysTemplate t on t.id = a.templateId
	;
	else 
		select a.templateId, t.headerId
		into v_templateId, v_headerId
		from nAnalys a 
		join nAnalysTemplate t on t.id = a.templateId
		where 
			a.byrow = p_byrow and a.bycolumn = p_bycolumn
		;
	end if;


	select 
		c.id as columnId
		, isnull(rc.sort, c.sort) as sort_token
		, c.name as columnName
		, name_ru as nameRu
		, isnull(rc.align, c.align) as align
    	, isnull(rc.hidden, c.hidden) as hidden
		, headType
		, isnull(us.width, c.width) as columnWidth
		, c.format
		, us.managId
	from nAnalysTemplate t
	join nResultHeader h on t.headerId = h.id
	join nResultColumns rc on rc.headerId = t.headerId
	join nResultColumnDef c on c.id = rc.columnId and (c.headType = p_headType or p_headType = 0)
	left join nHeaderColumnSelected us on us.managId = p_managId and us.templateId = t.id and us.columnId = c.id
	where t.id = v_templateId
	order by headType, sort_token;


end