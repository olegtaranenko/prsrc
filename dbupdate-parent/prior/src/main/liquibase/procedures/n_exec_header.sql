if exists (select 1 from sysprocedure where proc_name = 'n_exec_header') then
	drop procedure n_exec_header;
end if;


CREATE procedure n_exec_header (
	  p_filterId    integer
)
begin
	declare v_column_token  varchar(32);
	declare v_parent_token  varchar(32);
	declare v_proc_token    varchar(32);
	declare v_sub_token     varchar(32);
	declare v_sql           long varchar;
	declare v_begin         date;
	declare v_end           date;
	declare v_sqlHeader     varchar(254);



	create table #regions     (regionId     integer, isActive integer null);
	create table #materials   (klassId      integer, isActive integer null);
	create table #oborudItems (oborudItemId integer, isActive integer null);
	create table #noOboruds   (noOborud     integer, isActive integer null);
	create table #client      (clientId     integer, isActive integer null);


	for x as xd dynamic scroll cursor for
		call n_filter_params(p_filterid)
	do
		if r_itemType = 'filterPeriod' then
			if r_paramType = 'periodStart' then
				set v_begin = convert(date, r_charValue, 104);
			end if;
			if r_paramType = 'periodEnd' then
				set v_end = convert(date, r_charValue, 104);
			end if;
		else
			set v_sql = 'insert into #' + r_itemType;
			if r_paramClass = 'ids' then
				set v_sql = v_sql + '( ' + r_paramType + ', isActive)';
			end if;
    
			set v_sql = v_sql + ' values (';
			if r_intValue is not null then
				set v_sql = v_sql + convert(varchar(20), r_intValue) + ', ' + convert(varchar(20), r_isActive);
			elseif r_charValue is not null then
				set v_sql = v_sql + '''' + r_charValue + '''' + ', ' + convert(varchar(20), r_isActive);
			else
				set v_sql = v_sql + convert(varchar(20), r_isActive) + ', null';
			end if;
			set v_sql = v_sql + ')';
    
			message v_sql to client;
			execute immediate v_sql;
		end if;
	end for;



	select c.name, p.name, t.sqlHeader
	into v_column_token, v_parent_token, v_sqlHeader
	from nAnalys a
	join nAnalysCategory c on c.id = a.bycolumn
	left join nAnalysCategory p on p.id = c.parentId
	join nAnalysTemplate t on a.templateId = t.id
	join nFilter f on f.id = p_filterId and f.byrowid = a.byrow and f.bycolumnid = a.bycolumn
	;

	if v_parent_token is not null then
		set v_proc_token = v_parent_token;
		set v_sub_token = v_column_token;
	else 		
		set v_proc_token = v_column_token;
		set v_sub_token = null;
	end if;

	create table #periods (
		  periodId      int default autoincrement
		, klassId       integer       null
		, ventureId     integer       null
		, label         varchar(32)
		, st            date          null
		, en            date          null
		, year          integer       null
	);
	
	set v_sql = v_sqlHeader; 

	execute immediate v_sql;

	select * from #periods order by 1;

end;

