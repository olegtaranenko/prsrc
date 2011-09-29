if exists (select 1 from sysprocedure where proc_name = 'n_sell_item_cost') then
	drop procedure n_sell_item_cost;
end if;


CREATE PROCEDURE n_sell_item_cost (
	  p_begin         date
	, p_end           date
	, do_calc         integer
)
begin
-- перед вызовом этой процедуры должны быть:
-- а) создана таблица #sale_item. ѕри этом она должна быть пустой
-- б) создана таблица #orders. 
-- в) #orders должна быть заполнена заказами, подпадающими под другие фильтры
--  	как то, даты начала, конца, фирма, регион и так далее

-- Ќа выходе получаем набитую отдельными номенклатурами таблицу #sale_time с ценами, по которым они продались.
-- ƒаже если она продавалсь не по отдельности, а  в составе предмета, то ее цена вычисл€етс€ как часть от общей цены
-- путем сравнивани€ цен отделных номенклатур, вход€щих в предмет на момент составлени€ заказа (orders.inDate)
/*
	-- об€зательные пол€ таблиц
	create table #orders(
		  numorder   integer primary key
		, orderPaid       float	null
		, orderOrdered    float null
		, indate     date
		, periodid   integer	null
		, firmId     integer
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
		,periodid    integer null
		,priceToDate float null
		,quantEd     float null
	);
    
	-- естественно можно добавл€ть другие столбцы
*/
	insert into #sale_item (
		 numorder
		,nomnom
		,prId
		,prExt
		,materialQty
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
		, o.inDate
		, o.firmId
		, n.klassid
		, o.periodId
	from #orders o 
	join itemSellOrde i on o.numorder = i.numorder
	join sguidenomenk n on i.nomnom = n.nomnom
	where 
			o.indate >= p_begin and o.inDate < p_end 
		and exists (select 1 from #materials m where n.klassid = m.klassid)
	;
    
    if (do_calc = 1) then
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
			and p.cost >= 0.01
		group by p.nomnom
		;
	
		
		update #sale_item
		set priceToDate = p.cost
		from sPriceHistory p
			,#cost_to_date as ptd
		where 
			p.nomnom = #sale_item.nomnom
		and p.change_date = ptd.change_date
		and ptd.nomnom = #sale_item.nomnom
		;
	
		update #sale_item
		set priceToDate =  p.cost
		from sGuideNomenk p
		where p.nomnom = #sale_item.nomnom
			and #sale_item.priceToDate is null 
		;
	
		update	#sale_item 
		set cenaEd = k.cenaEd * n.perList
		from xPredmetyByNomenk k
			,sGuideNomenk n 
		where 
			k.nomnom = #sale_item.nomnom 
		and k.numorder = #sale_item.numorder
		and #sale_item.prId is null
		and n.nomnom = k.nomnom
		;
	
		update	#sale_item 
		set quantEd = i.quantEd / n.perList
		from itemBranOrde i
		join sGuideNomenk n on i.nomnom = n.nomnom
		where 
			i.numorder = #sale_item.numorder
		and i.prId = #sale_item.prId
		and i.prExt = #sale_item.prExt
		and i.nomnom = #sale_item.nomnom
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
	end if;
end