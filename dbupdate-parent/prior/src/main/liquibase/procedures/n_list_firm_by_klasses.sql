if exists (select 1 from sysprocedure where proc_name = 'n_list_firm_by_klasses') then
	drop procedure n_list_firm_by_klasses;
end if;


CREATE procedure n_list_firm_by_klasses (
	  p_begin         date
	, p_end           date
	, p_period_type   varchar(20) -- p_sub_token
	, p_rowId         integer
	, p_columnId      integer
)
begin

	declare v_table_name  varchar(64);
	declare v_ord_table   varchar(64);

	declare v_firmId      integer;
	declare v_klassId     integer;

 	message 'p_begin       = ', p_begin       to client;
	message 'p_end         = ', p_end         to client;
	message 'p_period_type = ', p_period_type to client;
	message 'p_rowId       = ', p_rowId       to client;
	message 'p_columnId    = ', p_columnId    to client;



	create table #sale_item (
		 numorder    integer
		,nomnom      varchar(20)
		,prId        integer null
		,prExt       integer null
		,materialQty float null
		,cenaEd      float null
		,inDate      date
		,firmId      integer
		,klassid     integer
		,periodid    integer null
		,priceToDate float null
		,quantEd     float null
	);

	set v_table_name = 'sGuideKlass';
	set v_ord_table = get_tmp_ord_table_name(v_table_name);
	execute immediate get_tmp_ord_create_sql(v_ord_table); -- #sGuideKlass_ord


	call n_internal_klasses (p_begin, p_end, v_table_name, p_rowId, p_columnId, 1);

	set v_firmId = p_rowId;

	
	if isnull(v_firmId, 0) = 0 then
		insert into #results (
			  label
			, materialQty
			, materialSaled
			, firm
			, region
			, regionid
			, periodid
			, firmId
			, oborud
		) select 
			  p.label
			, i.materialQty         -- к-во проданных единиц по выбранным материалам (шт, листов и т.д.)
			, i.materialSaled    	-- сумма по выбраннм материалам
			, f.name                -- фирма
			, r.region
			, r.regionid
			, p.klassid
			, f.firmId
			, f.tools as oborud
		from #periods p 
		join (

			select 
				sum(si.cenaEd * si.materialQty) as materialSaled
				, sum(si.materialQty) as materialQty
				, si.firmid
				, si.klassId
			from #sale_item si
			group by 
				si.firmid, si.klassId
		) i on 
			i.klassId = p.klassId
		join firmguide f on f.firmid = i.firmid
		join bayregion r on r.regionid = f.regionid
		;
	else
		insert into #results (
			  materialQty
			, materialSaled
			, indate
			, numorder
		)
		select 
			  i.materialQty
			, i.materialSaled
			, o.indate
			, o.numorder
		from bayorders o 
		join (
			select sum(si.cenaEd * si.materialQty) as materialSaled, sum(si.materialQty) as materialQty, numorder
			from #sale_item si
			group by si.numorder
		) i on 
			i.numorder = o.numorder
		;

	end if;

end;

