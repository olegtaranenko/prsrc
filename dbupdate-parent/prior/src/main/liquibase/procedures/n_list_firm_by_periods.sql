if exists (select 1 from sysprocedure where proc_name = 'n_list_firm_by_periods') then
	drop procedure n_list_firm_by_periods;
end if;


CREATE PROCEDURE n_list_firm_by_periods (
	  p_begin         date
	, p_end           date
	, p_period_type   varchar(20) -- p_sub_token
	, p_rowId         integer
	, p_columnId      integer
)
begin

	declare v_region_flag    integer;
	declare v_material_flag  integer;
	declare v_bayStatus_flag integer;
	declare v_tool_flag      integer;
	declare v_detail         integer;
	declare v_detail_fine    integer;
	declare v_begin          date;
	declare v_end            date;
	declare v_firmId         integer;

	set v_firmId = p_rowId;

	set v_detail      = 0;
	set v_detail_fine = 0;

	message 'p_begin       ', p_begin       to client;
	message 'p_end         ', p_end         to client;
	message 'p_period_type ', p_period_type to client;
	message 'p_rowId       ', p_rowId       to client;
	message 'p_columnId    ', p_columnId    to client;


	if isnull(p_rowId, 0) != 0 then
		set v_detail = 1;
		if isnull(p_columnId, 0) != 0 then
			set v_detail_fine = 1;
		end if;
	end if;

	select count(*) into v_region_flag   	from #regions   where isActive = 1;
	select count(*) into v_material_flag 	from #materials where isActive = 1;

	select toolId into v_tool_flag  from #tool;
	if v_tool_flag is null then
		set v_tool_flag = 0
	end if;


	
	select bayStatusId into v_bayStatus_flag  from #bayStatus;

	if v_bayStatus_flag is null then
		set v_bayStatus_flag = 0
	end if;


	call n_fill_periods(p_begin, p_end, p_period_type, p_columnId);

	set v_begin = p_begin;
	set v_end   = p_end;
	if v_detail_fine = 1 then 
		select st, en 
		into v_begin, v_end
		from #periods where periodId = p_columnId;
	end if;


	create table #orders(
		  numorder   integer primary key
		, orderPaid       float	null
		, orderOrdered    float null
		, indate     date
		, periodid   integer	null
		, firmId     integer
		, hasMaterial	integer	not null default 0
	);

	
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
		,periodid    integer
		,priceToDate float null
		,quantEd     float null
	);
    
	create table #firm_has_mat (
		firmId      integer primary key
	);


	insert into #orders (
		numorder, 
		indate, 
		firmId, 
		orderPaid, 
		orderOrdered
	)
	select 
		o.numorder, 
		o.indate, 
		o.firmId, 
		o.paid, 
		o.ordered
	from 
		bayorders o
--	join orderSellOrde s on s.numorder = o.numorder
	where 
			o.indate >= isnull(v_begin, o.inDate) and (v_end is null or o.inDate < v_end)
		and (v_region_flag = 0 or exists (select 1 from #regions r, bayguidefirms f where f.firmid = o.firmid and r.regionid = f.regionid))
		and (v_detail = 0 or o.firmId = v_firmId)
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


--select * from #orders;

	update #orders s set s.periodId = p.periodId
	from #periods p 
	where 
		s.indate >= p.st and s.inDate < p.en
	;

	
	if v_material_flag > 0 then
    
		insert into #sale_item (
			 numorder
			,nomnom
			,prId
			,prExt
			,materialQty
--			,sm
			,inDate
			,firmId      
			,klassid
			,periodId
		)
		select
			  o.numorder as numorder
			, i.nomnom
			, i.prId
			, i.prExt
			, i.quant / n.perlist as materialQty
--			, (i.quant) * i.cenaEd as sm
			, o.inDate
			, o.firmId
			, n.klassid
			, o.periodId
		from #orders o 
		join itemSellOrde i on o.numorder = i.numorder
		join sguidenomenk n on i.nomnom = n.nomnom
		where 
				o.indate >= v_begin and o.inDate < v_end 
			and exists (select 1 from #materials m where n.klassid = m.klassid)
		;
    

		update #orders s set s.hasMaterial = 1
		where 
			exists (select 1 from #sale_item i where i.numorder = s.numorder)
		;

		insert into #firm_has_mat
		select distinct firmId 
		from #orders s
		where s.hasMaterial = 1
		;


		create table #cost_to_date(
			nomnom varchar(20)
--			,cost float null
			,change_date datetime
		)
		;

		insert into #cost_to_date(
			nomnom
--			,cost
			,change_date
		)
		select 
			p.nomnom 
			,min(p.change_date)
		from 
			sPriceHistory p
		join 
			#sale_item si on si.nomnom = p.nomnom
		where
			p.change_date >= si.inDate
		group by p.nomnom
		;

		
		update #sale_item
		set priceToDate = p.cost
		from sPriceHistory p
			,#cost_to_date as ptd
		where 
			p.nomnom = #sale_item.nomnom
		and p.change_date = ptd.change_date
		;

		update #sale_item
		set priceToDate =  p.cost
		from sGuideNomenk p
		where p.nomnom = #sale_item.nomnom
			and #sale_item.priceToDate is null
		;

		update	#sale_item 
		set cenaEd = k.cenaEd
		from xPredmetyByNomenk k
		where 
			k.nomnom = #sale_item.nomnom 
		and k.numorder = #sale_item.numorder
		and #sale_item.prId is null
		;

		update	#sale_item 
		set quantEd = i.quantEd
		from itemBranOrde i
		where 
			i.numorder = #sale_item.numorder
		and i.prId = #sale_item.prId
		and i.prExt = #sale_item.prExt
		and #sale_item.prId is not null
		;


		create table #k_cost (
			numorder int
			,prId int
			,prExt int
			,k_cost float null
		)
		;		


		insert into #k_cost(
			numorder
			,prId
			,prExt
			,k_cost
		)
		select
			i.numorder 
			,i.prId
			,i.prExt
			,cur.total / i.cenaEd --/  as k_cost
		from
		(
			select numorder, prId, prExt, sum(priceToDate * quantEd) as total
			from #sale_item si
			where prId is not null
			group by numorder, prId, prExt
		) as cur
		,xPredmetyByIzdelia i 
		where cur.numorder = i.numorder and cur.prId = i.prId and cur.prExt = i.prExt
			and isnull(i.cenaEd, 0) > 0
		;


		update	#sale_item 
		set cenaEd = priceToDate / k.k_cost
		from #k_cost k
		where 
			k.numorder = #sale_item.numorder
		and k.prId = #sale_item.prId
		and k.prExt = #sale_item.prExt
		and k.k_cost > 0
		;


	else
		insert into #firm_has_mat
		select distinct firmId 
		from #orders s
	end if;



	if v_detail = 0 then
		insert into #results (
			  label
			, year
			, orderQty
			, orderPaid
			, orderOrdered
			, orderMatQty
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
			, o.orderMatQty         -- к-во заказов, в которых присутствовали выбранные материалы
			, i.materialSaled    	-- сумма по выбранным материалам (если нет фильтрации - совпадает общим количеством)
			, f.name                -- фирма
			, r.region
			, r.regionid
			, p.periodid
			, o.firmId
			, f.tools as oborud
		from #periods p 
		join (
			select sum(isnull(orderPaid, 0)) as orderPaid
				, count(*) as orderQty, firmId, periodId
				, sum(orderOrdered) as orderOrdered
				, sum(hasMaterial) as orderMatQty
			from #orders
			group by firmId, periodId
		) o on 
			o.periodid = p.periodId
		join #firm_has_mat fm on fm.firmid  = o.firmid
		join bayguidefirms f  on f.firmid   = o.firmid
		join bayregion     r  on r.regionid = f.regionid
		left join (
			select 
				sum(isnull(si.cenaEd, 0) * isnull(si.materialQty, 0)) as materialSaled
				, si.firmid
				, si.periodId
			from 
				#sale_item si
			group by 
				firmid
				, periodId
		) i on 
			o.firmId = i.firmId and o.periodId = i.periodId
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
		from #orders s
		left join (
			select sum(isnull(si.cenaEd, 0) * isnull(si.materialQty, 0)) as materialSaled
			, sum(si.materialQty) as materialQty
			, si.numorder
			from #sale_item si
			group by si.numorder
		) i on 
			i.numorder = s.numorder
		;

	end if;

end;


