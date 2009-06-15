if exists (select 1 from sysprocedure where proc_name = 'n_list_climat_by_periods') then
	drop procedure n_list_climat_by_periods;
end if;


CREATE procedure n_list_climat_by_periods (
	  p_begin         date
	, p_end           date
	, p_period_type   varchar(20) -- p_sub_token
	, p_rowId         integer
	, p_columnId      integer
)
begin
	declare v_region_flag integer;
	declare v_material_flag integer;
	declare v_oborud_flag integer;
	declare v_no_oborud_flag integer;

	declare v_detail      integer;
	declare v_detail_fine integer;

	declare v_begin       date;
	declare v_end         date;

	declare v_firmId      integer;

	set v_firmId = p_rowId;

	set v_detail      = 0;
	set v_detail_fine = 0;
	if isnull(p_rowId, 0) != 0 then
		set v_detail = 1;
		if isnull(p_columnId, 0) != 0 then
			set v_detail_fine = 1;
		end if;
	end if;

	select count(*) into v_region_flag   	from #regions   where isActive = 1;
	select count(*) into v_material_flag 	from #materials where isActive = 1;
	select count(*) into v_no_oborud_flag   from #noOboruds;
	set v_oborud_flag = 0;
	if v_no_oborud_flag = 0 then
		select count(*) into v_oborud_flag   from #oborudItems where isActive = 1;
	end if;


	call n_fill_periods(p_begin, p_end, p_period_type, p_columnId);

	set v_begin = p_begin;
	set v_end   = p_end;
	if v_detail_fine = 1 then 
		select st, en 
		into v_begin, v_end
		from #periods where periodId = p_columnId;
	end if;

	
	create table #sale_isum(
		  numorder   integer
		, orderPaid       float
		, orderOrdered    float
		, indate     date
		, periodid   integer
		, firmId     integer
	);


	insert into #sale_isum (
		numorder, indate, firmId, orderPaid, orderOrdered
	)
	select o.numorder, o.indate, o.firmId, o.paid, s.cena
	from 
		bayorders o
	join orderSellOrde s on s.numorder = o.numorder
	where 
			o.indate >= isnull(v_begin, o.inDate) and (v_end is null or o.inDate < v_end)
		and (v_region_flag = 0 or exists (select 1 from #regions r, bayguidefirms f where f.firmid = o.firmid and r.regionid = f.regionid))
		and (v_detail = 0 or o.firmId = v_firmId)
		and (v_oborud_flag = 0 
			or exists (
				select 1 from oborudKomplekt ok 
				, #oborudItems oi, bayGuideFirms f
				where ok.oborudId = f.oborudId and ok.oborudItemId = oi.oborudItemId and f.firmId = o.firmId
			)
		)
		and (v_no_oborud_flag = 0 
			or exists (
				select 1 from bayGuideFirms f
				where f.firmId = o.firmId and f.oborudId is null
			)
		)
	;


	update #sale_isum s set s.periodId = p.periodId
	from #periods p 
	where 
		s.indate >= p.st and s.inDate < p.en
	;

	
	create table #sale_item (
		 numorder    integer
		,nomnom      varchar(20)
		,materialQty         float
		,sm          float
		,inDate      date
		,firmId      integer
		,klassid     integer
		,periodid    integer
	);


	insert into #sale_item (
		 numorder
		,nomnom
		,materialQty
		,sm
		,inDate
		,firmId      
		,klassid
		,periodId
	)
	select
		  o.numorder as numorder
		, i.nomnom
		, i.quant / n.perlist as materialQty
		, (i.quant / n.perlist) * i.cenaEd as sm
		, o.inDate
		, o.firmId
		, n.klassid
		, si.periodId
	from itemSellOrde i
	join bayorders o on o.numorder = i.numorder 
	join sguidenomenk n on i.nomnom = n.nomnom
	join #sale_isum si on si.numorder = i.numorder
	where 
			o.indate >= v_begin and o.inDate < v_end 
		and (v_material_flag = 0 or exists (select 1 from #materials m where n.klassid = m.klassid))
	;



	if v_detail = 0 then
		insert into #results (
			  label
			, year
			, orderQty
			, orderPaid
			, orderOrdered
			, materialQty
			, materialSaled
			, firm
			, region
			, regionid
			, periodid
			, firmId
			, oborud
		)
		select 
			  p.label
			, p.year
			, o.orderQty            -- число заказов за период
			, o.orderPaid           -- общий объем заказов (уе)
			, o.orderOrdered        -- общая сумма по заказам
			, i.materialQty         -- к-во проданных единиц по выбранным материалам (шт, листов и т.д.)
			, i.materialSaled    	-- сумма по выбраннм материалам
			, f.name                -- фирма
			, r.region
			, r.regionid
			, p.periodid
			, o.firmId
			, ob.oborud
		from #periods p 
		join (
			select sum(sm) as materialSaled, sum(materialQty) as materialQty, firmid, periodId
			from #sale_item
			group by firmid, periodId
		) i on 
			i.periodid = p.periodId
		join (
			select sum(isnull(orderPaid, 0)) as orderPaid, count(*) as orderQty, firmId, periodId, sum(orderOrdered) as orderOrdered
			from #sale_isum
			group by firmId, periodId
		) o on 
			o.firmId = i.firmId and o.periodId = i.periodId
		join bayguidefirms f on f.firmid = i.firmid
		join bayregion r on r.regionid = f.regionid
		left join guideOborud ob on ob.oborudId = f.oborudId
		;
	elseif v_detail = 1 then
		insert into #results (
			  orderPaid
			, orderOrdered
			, materialQty
			, materialSaled
			, indate
			, numorder
		)
		select 
			  s.orderPaid
			, s.orderOrdered
			, i.materialQty
			, i.materialSaled
			, s.indate
			, s.numorder
		from #sale_isum s 
		join (
			select sum(sm) as materialSaled, sum(materialQty) as materialQty, numorder
			from #sale_item
			group by numorder
		) i on 
			i.numorder = s.numorder
		;

	end if;
	
end;
