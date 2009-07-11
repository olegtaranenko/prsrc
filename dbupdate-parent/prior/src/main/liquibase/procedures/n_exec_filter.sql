if exists (select 1 from sysprocedure where proc_name = 'n_exec_filter') then
	drop function n_exec_filter;
end if;


CREATE procedure n_exec_filter (
	  p_filterId    integer
	, p_rowId       integer default 0
	, p_columnId    integer default 0
	, p_showFirst   integer default 1
	, p_showLast    integer default 1
)
begin
	declare v_sql long varchar;
	declare groupByRows    varchar(64);
	declare groupByColumns varchar(64);
	declare v_periodType     varchar(64);
	declare v_begin       date;
	declare v_end         date;
	declare v_proc_name   varchar(128);

	declare v_column_token  varchar(32);
	declare v_parent_token  varchar(32);
	declare v_proc_token    varchar(32);
	declare v_sub_token     varchar(32);
	declare v_sqlFunction   varchar(127);

	declare v_firstVisit    varchar(1);


	select r.name, c.name, p.name, t.sqlFunction
	into groupByRows, v_column_token, v_parent_token, v_sqlFunction
	from nAnalys a
	join nAnalysCategory r on r.id = a.byrow
	join nAnalysCategory c on c.id = a.bycolumn
	left join nAnalysCategory p on p.id = c.parentId
	join nAnalysTemplate t on a.templateId = t.id
	join nFilter f on f.id = p_filterId and f.byrowid = a.byrow and f.bycolumnid = a.bycolumn
	;


	if v_parent_token is not null then
		set groupByColumns = v_parent_token;
		set v_sub_token = v_column_token;
	else 		
		set groupByColumns = v_column_token;
		set v_sub_token = null;
	end if;


	create table #regions (regionId integer, isActive integer null);
	create table #materials (klassId integer, isActive integer null);
	create table #oborudItems (oborudItemId integer, isActive integer null);
	create table #noOboruds (noOborud integer, isActive integer null);
	create table #client      (clientId integer, isActive integer null);
	
	for x as xc dynamic scroll cursor for
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
    
--			message v_sql to client;
			execute immediate v_sql;
		end if;
	end for;


	create table #results (
		  label         varchar(64)  null
		, year          integer      null
		, orderQty      integer      null
		, orderPaid     float        null
		, orderOrdered  float        null
		, materialQty   float        null
		, materialSaled float        null
		, firm          varchar(512) null
		, region        varchar(256) null
		, regionid      integer      null
		, periodid      integer      null
		, firmId        integer      null
		, inDate        date         null
		, numorder      integer      null
		, oborud        varchar(32)  null
		, nomnom        varchar(20)  null
		, nomname       varchar(128) null
		, edizm         varchar(10)  null
		, cena          float        null
		, matInQty      float        null
		, matInTurn     float        null
		, matOutTurn    float        null
		, matOutQty     float        null
		, sumOut        float        null
		, orderMatQty   float        null
	);

	create table #periods (
		  periodId      integer      default autoincrement
		, klassId       integer      null
		, ventureId     integer      null
		, label         varchar(32)  null
		, st            date         null
		, en            date         null
		, year          integer      null
	);


	call n_default_period(v_begin, v_end, p_filterId);

	set v_proc_name = 'n_list_' + groupByRows + '_by_' + groupByColumns;

	execute immediate 'call ' + v_proc_name + '(v_begin, v_end, v_sub_token, p_rowId, p_columnId)';


	set v_firstVisit = n_get_booting_param(p_filterId, 'firstVisit');
	if v_firstVisit = '1' then
		create table #firm_besuch (firmId integer, firstVisit date, lastVisit date);
		if p_showFirst is not null or p_showLast is not null then
        
			insert into #firm_besuch (firmId, firstVisit, lastVisit) 
			select firmId, min(o.inDate), max(o.inDate)
			from bayOrders o
			where 
				exists (select 1 from #results r where r.firmId = o.firmId)
			group by firmId;
		end if;
	end if;


	-- достаточно корявое решение. Нужно бы перенести выдачу результата в функцию подготавливающую данные.
	-- 
	if p_rowId = 0 and p_columnId = 0 then
		-- выдача резалт-сета с полями первого и последнего посещения.
		if v_firstVisit = '1' then
			select r.*, b.firstVisit, b.lastVisit
			from #results r
			left join #firm_besuch b on b.firmId = r.firmId
			order by r.firm, r.firmid, r.periodid;
		else
			select r.*
			from #results r
			order by r.nomnom;
		end if;
	elseif p_rowId != 0 then
		select r.* 
		from #results r
		order by r.numorder;
	end if;
end;


