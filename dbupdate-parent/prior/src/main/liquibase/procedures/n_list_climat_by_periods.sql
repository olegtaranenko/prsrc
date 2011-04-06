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

	declare v_detail      integer;
	declare v_detail_fine integer;

	declare v_begin       date;
	declare v_end         date;

	declare v_clientId      integer;

	set v_detail      = 0;
	set v_detail_fine = 0;
	if isnull(p_rowId, 0) != 0 then
		set v_detail = 1;
		if isnull(p_columnId, 0) != 0 then
			set v_detail_fine = 1;
		end if;
	end if;

	select clientId into v_clientId from #client;


	call n_fill_periods(p_begin, p_end, p_period_type, p_columnId);

	set v_begin = p_begin;
	set v_end   = p_end;
	if v_detail_fine = 1 then 
		select st, en 
		into v_begin, v_end
		from #periods where periodId = p_columnId;
	end if;

	
	create table #sale_item (
		  numorder integer
		, nomnom   varchar(20)
		, indate   date
		, quant    float
		, cenaEd   float          null
		, periodid integer        null
		, costEd float          null
		, edIzm    varchar(20)    null
		, itemName varchar (128) null
	);


	insert into #sale_item (
		numorder, nomnom, indate
		, quant, cenaEd
		, edIzm, costEd
		, itemName
	)
	select 
		o.numorder, s.nomnom, o.indate
		, s.quant / n.perlist, s.cenaEd * n.perlist
		, n.ed_izmer2, n.cena_W
		, trim(n.cod + ' ' + n.nomname + ' ' + n.size)             -- название
	from 
		orders o
	join itemBareOrde s on s.numorder = o.numorder
	join sguidenomenk n on s.nomnom = n.nomnom
	where 
			o.indate >= isnull(v_begin, o.inDate) and (v_end is null or o.inDate < v_end)
		and o.firmId = v_clientId
	;

	insert into #sale_item (
		numorder, nomnom, indate
		, quant, cenaEd
		, edIzm, costEd
		, itemName
	)
	select 
		o.numorder, gp.prName, o.indate
		, s.quant , s.cenaEd
		, 'Изд.', gp.Cena4
		, trim(gp.prName + ' ' + gp.prDescript + ' ' + gp.prSize)             -- название
	from 
		orders o
	join itemIzdeOrde s on s.numorder = o.numorder
	join sguideproducts gp on s.prId = gp.prId
	where 
			o.indate >= isnull(v_begin, o.inDate) and (v_end is null or o.inDate < v_end)
		and o.firmId = v_clientId
	;



	update #sale_item s set s.periodId = p.periodId
	from #periods p 
	where 
		s.indate >= p.st and s.inDate < p.en
	;

	
	create table #sale_isum (
		  nomnom      varchar(20)
		, materialQty float
		, materialSm  float          null
		, periodid    integer        null
		, costEd      float          null
		, edIzm       varchar(20)    null
		, itemName varchar (128) null
	);


	insert into #sale_isum (
		  nomnom
		, materialQty
		, materialSm
		, periodId
		, costEd, edIzm
		, itemName
	) select
		  i.nomnom
		, sum(i.quant)
		, sum(i.quant * i.cenaEd)
		, i.periodId
		, i.costEd, i.edIzm
		, i.itemName
	from  #sale_item i
	group by 
		i.periodId
		, i.nomnom
		, i.costEd
		, i.edIzm
		, i.itemName
	;

	if v_detail = 0 then
		insert into #results (
			  label
			, year
			, materialQty
			, materialSaled
			, nomnom
			, nomname
			, edizm
			, cena
			, periodId
		)
		select 
			  p.label
			, p.year
			, i.materialQty         -- к-во проданных единиц по выбранным материалам (шт, листов и т.д.)
			, i.materialSm    	-- сумма по выбраннм материалам
			, i.nomnom              -- номер номенклатуры
			, i.itemName
			, i.edIzm
			, i.costEd
			, i.periodid
		from #sale_isum i 
		join #periods   p on i.periodid = p.periodId
		;
	end if;

end;
