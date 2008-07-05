/*

	ИТОГО сухой остаток.
	Нужно обеспечить вывод номенклатуры единым списком по всем направлениям деятельности предприятия (пр-во, продажи)

	Префиксы (по типу списка):
		item:		Однократное вхождение позиции в список. Пример единчная отгрузка позиции, заказанное изделие,
					или позиция ном-ры, входящая через изделие.
		isum:		Суммирование однотипных вхождений номенктуры. Пример: сумма отгруженной позиции номенклатуры заказа, если отгрузок несколько
		order:		Суммирование по заказу. Пример: общая себестоимсо заказа.


	Корни по направлению деятельнсти:                                                     Значения Дискриминатора
		Prod:		от Production. Производство - на своей номенклатуре                   p  1,2.3
		Serv:		от Service.    Производство - услуги, давальческий материал.          u  4
		Sell:		от Sells.      Продажи                                                b  8
		Flaw:       от Flaw.       Брак                                                   f  -
		Bran:		от Branch.     По всем видам деятельности
	
		

		
	Корни по типу предметов, составляющих заказ:                                          Значения Дискриминатора          
		Ware:	Если относится к готовым изделиям, состоящим из отдельных номенклатуры    p,w  1
		Sepa:	Если относится только к отдельной номенклатуре                            n,w  2
		Vari:	Относится к переменной части вариантной номенклатуры                      p,w
		Fixe:	Относится к фиксированной части вариантной номенклатуры или               p,w
				просто к невариантному изделию                                                 
		Wall:	от Wares All. По всем типам предметов.
		Rsrv:   от Reserved. зарезервированная номенклатура
			

	Cуффиксы:
		Orde	сокращение от Ordered:	заказанно клиентом для заказа
		Requ	сокращение от Requested: затребованно производством для выполнения заказа
		Rele	сокращение от Released:	отпущено со склада для производства
		Prep	сокращение от Prepared:	готово к отгрузке. (Пока не требуется, в 
					перспективе для более тщательного учета.)
		Proc	сокращение от Processed:"незавершека"
		Ship	сокращение от Shipped:	отгруженно клиенту для исполнения заказа

 	Список аттрибутов должен быть более менее стандартный и в него входят:
 		numorder:	номер заказа		
 				
		nomnom:		номер номенклатуры  
				
		quant:		кол-во номенклатуры в производственных ед. 
				
		date1:		дата1 (под вопросм)  - первая дата, относящаяся к операции, если операция занимает больше чем один день
					дата операции. Пример - дата выписки заказа
		 		
		date2:		дата2 (под вопросом) - последняя дата, относящаяся к операции, если операция занимает больше чем один день
					null, если это не применимо в контексте.
					в точности равна дате1, если операции совершились в одном дне.
					не может быть раньше чем дата1.
				
		costEd:		затраты (себестоимость) номенклатуры на дату1 на НЕ ПРОИЗВОДСТВЕННУЮ единицу номенклатуры (лист и т.д.)
					null, если это не применимо в контексте.
				
		cenaEd:		цена (отпускная) номенклатуры, null, если это не применимо в контексте.
					относится к НЕ ПРОИЗВОДСТВЕННОЙ единице номенклатуры (лист и т.д.)	
				
		ventureId:	предприятия, через которое проходит заказ. Null - либо не применимо, или старый заказ, до Интеграции.

	Список дополнительных атрибутов может меняться в зависимости от того, где вьюха будет применяться.
*/






if exists (select 1 from sysviews where viewname = 'itemWareFixe' and vcreator = 'dba') then
	drop view itemWareFixe;
end if;


create view itemWareFixe
-- список фиксированной части готовых изделий с вариантной номенклатурой.
as 
select * from 
sproducts p 
where not exists (select 1 from sguidevariant gv where p.productid = gv.productid and p.xgroup = gv.xgroup and not (gv.xgroup = '' or (gv.xgroup != '' and gv.c = 1)))
;



--=====================================================
--	Номенклатура заказанная 
-- 		Orde	сокращение от Ordered:		заказанно клиентом для заказа
--=====================================================

if exists (select 1 from sysviews where viewname = 'itemSellOrde' and vcreator = 'dba') then
	drop view itemSellOrde;
end if;


create view itemSellOrde (numorder, nomnom, quant, cenaEd, statusid)
as
select r.numdoc, r.nomnom, r.quantity as quant, r.intQuant, o.statusid
from 
sdmcrez r
join bayorders o on o.numorder = r.numdoc
;


if exists (select 1 from sysviews where viewname = 'isumSellOrde' and vcreator = 'dba') then
	drop view isumSellOrde;
end if;


create view isumSellOrde (numorder, nomnom, quant, cenaEd, statusid)
as
select i.numorder, i.nomnom, i.quant, i.cenaEd, i.statusid
from itemSellOrde i
;


if exists (select 1 from sysviews where viewname = 'itemProdOrde' and vcreator = 'dba') then
	drop view itemProdOrde;
end if;

create view itemProdOrde (numorder, nomnom, quant, cenaEd, prid, prext)
-- список номенклатуры, входящей в предметы заказа производства
-- заказы по продажам НЕ ВХОДЯТ!
-- Изделия (включая вариантные) разбираются на составные номенклатуры.
as 
select numorder, nomnom, round(fn.quantity * io.quant, 5), null, io.prid, io.prext
from xpredmetybyizdelia io 
join itemWareFixe fn on io.prid = fn.productid
	union all
select io.numorder, v.nomnom, round(p.quantity * io.quant, 5), null, io.prid, io.prext
from xpredmetybyizdelia io 
join xvariantnomenc v on v.numorder = io.numorder and v.prid = io.prid and v.prext = io.prext
join sproducts p on p.productid = io.prid and v.nomnom = p.nomnom
	union all
select po.numorder, po.nomnom, po.quant, po.cenaEd, null, null
from xpredmetybynomenk po
;




if exists (select 1 from sysviews where viewname = 'isumProdOrde' and vcreator = 'dba') then
	drop view isumProdOrde;
end if;

create view isumProdOrde(numorder, nomnom, quant, statusid)
as 
select i.numorder, i.nomnom, sum(i.quant), o.statusid
from itemProdOrde i
join orders o on o.numorder = i.numorder
group by i.numorder, i.nomnom, o.statusid
;


if exists (select 1 from sysviews where viewname = 'isumBranOrde' and vcreator = 'dba') then
	drop view isumBranOrde;
end if;

create view isumBranOrde (numorder, nomnom, quant, statusid, scope)
as
select numorder, nomnom, quant, statusid, 'p'
from isumProdOrde 
	union all
select numorder, nomnom, quant, statusid, 'b'
from itemSellOrde 
;


if exists (select 1 from sysviews where viewname = 'orderSellOrde' and vcreator = 'dba') then
	drop view orderSellOrde;
end if;

create view orderSellOrde (numorder, cena, statusid)
as
select i.numorder, sum(i.quant * i.cenaEd / n.perlist), i.statusid
from isumSellOrde i
join sguidenomenk n on i.nomnom = n.nomnom
group by i.numorder, i.statusid
;

--=====================================================


--=====================================================
--	Номенклатура: затребованно производством для выполнения заказа
--		Requ	сокращение от Requested:	
--=====================================================



if exists (select 1 from sysviews where viewname = 'itemProdRequ' and vcreator = 'dba') then
	drop view itemProdRequ;
end if;

create view itemProdRequ (numorder, nomnom, quant, statusid)
as 
select r.numdoc, r.nomnom, r.curquant, o.statusid
from sdmcrez r
join orders o on r.numdoc = o.numorder
where r.curquant > 0
;


if exists (select 1 from sysviews where viewname = 'isumProdRequ' and vcreator = 'dba') then
	drop view isumProdRequ;
end if;

create view isumProdRequ
as 
select *
from itemProdRequ r
;




if exists (select 1 from sysviews where viewname = 'itemSellRequ' and vcreator = 'dba') then
	drop view itemSellRequ;
end if;

create view itemSellRequ (numorder, nomnom, quant, statusid)
as 
select r.numdoc, r.nomnom, r.curquant * n.perlist, o.statusid
from sdmcrez r
join bayorders o on r.numdoc = o.numorder
join sguidenomenk n on r.nomnom = n.nomnom
where r.curquant > 0
;


if exists (select 1 from sysviews where viewname = 'isumSellRequ' and vcreator = 'dba') then
	drop view isumSellRequ;
end if;

create view isumSellRequ
as 
select *
from itemSellRequ r
;



if exists (select 1 from sysviews where viewname = 'itemBranRequ' and vcreator = 'dba') then
	drop view itemBranRequ;
end if;


create view itemBranRequ (numorder, nomnom, quant, scope, statusid)
as 
select r.numorder, r.nomnom, r.quant, 'p', r.statusid
from itemProdRequ r
		union all
select r.numorder, r.nomnom, r.quant, 'p', r.statusid
from itemSellRequ r
;


if exists (select 1 from sysviews where viewname = 'isumBranRequ' and vcreator = 'dba') then
	drop view isumBranRequ;
end if;

create view isumBranRequ
as 
select *
from itemBranRequ r
;



--=====================================================
--		Rele	сокращение от Released:	отпущено со склада для производства
--=====================================================


if exists (select 1 from sysviews where viewname = 'itemBranRele' and vcreator = 'dba') then
	drop view itemBranRele;
end if;

create view itemBranRele (numorder, nomnom, quant, scope, date1, statusid)
as 
select r.numdoc, r.nomnom, r.quant
, if o.numorder is not null then 'p' else if bo.numorder is not null then 'b' else '?' endif endif 
, d.xdate
, if o.numorder is not null then o.statusid else if bo.numorder is not null then bo.statusid else null endif endif 
from sdmc r
join sdocs d on d.numdoc = r.numdoc and d.numext = r.numext
left join orders o on r.numdoc = o.numorder
left join bayorders bo on r.numdoc = bo.numorder
where r.numext < 254
;



if exists (select 1 from sysviews where viewname = 'isumBranRele' and vcreator = 'dba') then
	drop view isumBranRele;
end if;

create view isumBranRele (numorder, nomnom, quant, scope, date1, date2, statusid)
as 
select r.numorder, r.nomnom, sum(r.quant) as quant
, r.scope
, min(r.date1) as date1, max(r.date1) as date2
, r.statusid
from itemBranRele r
where r.statusid < 6
group by r.numorder, r.nomnom, r.scope, r.statusid
;





--=====================================================
--		WareShip	отгружено по количеству
--=====================================================


if exists (select 1 from sysviews where viewname = 'itemWareShip' and vcreator = 'dba') then
	drop view itemWareShip;
end if;

create view itemWareShip (outdate, prId, prExt, numorder, nomnom, quant)
-- список номенклатуры, входящие в отгруженные готовые изделия.
-- несколько отгрузок порождают несколько строк с одинаковыми заказ-изделие-номенклатура.
-- учитываются изделия как с фиксированной, так и с вариантной номенклатурой.
as 
select outdate, io.prId, io.prExt, numorder, nomnom, round(fn.quantity, 5) as quant
from xpredmetybyizdeliaout io 
join itemWareFixe fn on io.prid = fn.productid
	union all
select outdate, io.prId, io.prExt, io.numorder, v.nomnom, round(p.quantity, 5) as quant
from xpredmetybyizdeliaout io 
join xvariantnomenc v on v.numorder = io.numorder and v.prid = io.prid and v.prext = io.prext
join sproducts p on p.productid = io.prid and v.nomnom = p.nomnom
;


if exists (select 1 from sysviews where viewname = 'orderWareShip' and vcreator = 'dba') then
	drop view orderWareShip;
end if;

create view orderWareShip (outdate, prId, prExt, numorder, costEd)
-- себестоимость готовых изделий входящие в заказы, которые уже отгруженны.
as 
select outdate, prid, prext, numorder, sum(round(n.cost * io.quant / n.perlist, 2))
from itemWareShip io 
join sGuideNomenk n on io.nomnom = n.nomnom
group by outdate, prid, prext, numorder
;



if exists (select 1 from sysviews where viewname = 'itemWallShip' and vcreator = 'dba') then
	drop view itemWallShip;
end if;


create view itemWallShip (outdate, numorder, type, prId, prExt, prNomnom, cenaEd, quant, costEd, firmName, ventureId, statusid)
-- единым списком выводится суммы по затратам и реализации заказов, которые можно рассматривать как отгруженные
-- по всем направлениям деятельности: производственные (изделия и отдельная ном-ра), услуги, продажи.
-- используется в отчете Реализация Товаров
as 
select 
	po.outdate, po.numorder, 1, po.prId, po.prExt, null, p.cenaEd, po.quant, io.costEd
	, f.name, o.ventureid, o.statusid
from xpredmetybyizdeliaout po
join xpredmetybyizdelia p on p.numorder = po.numorder and p.prid = po.prid and p.prext = po.prext
join orderWareShip io on po.outdate = io.outdate and io.numorder = po.numorder and io.prid = po.prid and io.prext = po.prext
join orders o on o.numorder = po.numorder and o.numorder = p.numorder
join guidefirms f on f.firmid = o.firmid
	union all 
select po.outdate, po.numorder, 2, null, null, po.nomnom, p.cenaEd, po.quant, round(round(n.cost, 2) / n.perlist, 2) as costEd
	, f.name, o.ventureid, o.statusid
from xpredmetybynomenkout po
join xpredmetybynomenk p on p.numorder = po.numorder and p.nomnom = po.nomnom
join sguidenomenk n on n.nomnom = po.nomnom and n.nomnom = p.nomnom
join orders o on o.numorder = po.numorder and o.numorder = p.numorder
join guidefirms f on f.firmid = o.firmid
	union all 
select u.outdate, u.numorder, 4, null, null, null, 1.0, u.quant, null
	, f.name, o.ventureid, o.statusid
from xuslugout u
join orders o on o.numorder = u.numorder 
join guidefirms f on f.firmid = o.firmid
	union all 
select po.outDate, o.numOrder, 8, null, null, po.nomnom, r.intQuant AS cenaed, po.quant, n.cost as costEd
	, f.name, o.ventureid, o.statusid
from bayorders o
--join sDocs d on d.numDoc = o.numOrder 
join sDMCrez r on r.numDoc = o.numorder
join baynomenkout po on po.numorder = o.numorder and po.nomnom = r.nomnom
join sguidenomenk n on n.nomnom = po.nomnom
join bayguidefirms f on f.firmid = o.firmid
;




if exists (select 1 from sysviews where viewname = 'orderWallShip' and vcreator = 'dba') then
	drop view orderWallShip;
end if;


create view orderWallShip (outdate, numorder, type, cenaTotal, costTotal, firmname, ventureid)
-- список, группирующий все отгруженное по заказам.
as 
select outdate, numorder, sum(distinct(type)), sum(isnull(round(quant * cenaEd , 2), 0)), sum(isnull(round(quant * costEd, 2), 0)), firmname, ventureid
from itemWallShip po
group by outdate, numorder, firmname, ventureid;





--=====================================================
--	Ship сокращение от Ship: отгружено по номенклатуре
--=====================================================

if exists (select 1 from sysviews where viewname = 'itemSellShip' and vcreator = 'dba') then
	drop view itemSellShip;
end if;

create view itemSellShip (numorder, nomnom, quant, date1)
-- Отгруженная номенклатура единым списком, включая продажи.
-- на каждую отгрузку или на каждое вхождение номенклатуры заказа через изделия - своя строчка
as 
select io.numorder, io.nomnom, io.quant * n.perlist, outdate
from baynomenkout io
join sguidenomenk n on n.nomnom = io.nomnom
;

if exists (select 1 from sysviews where viewname = 'isumSellShip' and vcreator = 'dba') then
	drop view isumSellShip;
end if;


create view isumSellShip (numorder, nomnom, quant, date1, date2)
as 
select numorder, nomnom, sum(quant) as quant, min(date1), max(date1)
from itemSellShip
group by
	numorder, nomnom
;





if exists (select 1 from sysviews where viewname = 'orderSellShip' and vcreator = 'dba') then
	drop view orderSellShip;
end if;



create view orderSellShip (numorder, cenaTotal, statusid)
as 
	select o.numorder, sum(r.intQuant * po.quant) as cenaTotal, o.statusid
	from bayorders o
	join sDMCrez r on r.numDoc = o.numorder
	join baynomenkout po on po.numorder = o.numorder and po.nomnom = r.nomnom
	group by o.numorder, o.statusid
;




if exists (select 1 from sysviews where viewname = 'itemProdShip' and vcreator = 'dba') then
	drop view itemProdShip;
end if;



create view itemProdShip (numorder, nomnom, quant, date1)
-- Отгруженная номенклатура единым списком, включая продажи.
-- на каждую отгрузку или на каждое вхождение номенклатуры заказа через изделия - своя строчка
as 
select numorder, nomnom, round(fn.quantity * io.quant, 5) as quant, io.outdate
-- отгрзка ФИКСИРОВАННОЙ части вариантного изделия, а также номенклатуры невариантного изделия.
from xpredmetybyizdeliaout io 
join itemWareFixe fn on io.prid = fn.productid
	union all
select io.numorder, v.nomnom, round(p.quantity * io.quant, 5) as quant, io.outdate
-- отгрзка ПЕРЕМЕННОЙ части вариантного изделия
from xpredmetybyizdeliaout io 
join xvariantnomenc v on v.numorder = io.numorder and v.prid = io.prid and v.prext = io.prext
join sproducts p on p.productid = io.prid and v.nomnom = p.nomnom
    union all
select io.numorder, io.nomnom, io.quant, io.outdate
from xpredmetybynomenkout io
;


if exists (select 1 from sysviews where viewname = 'isumProdShip' and vcreator = 'dba') then
	drop view isumProdShip;
end if;


create view isumProdShip (numorder, nomnom, quant, date1, date2)
as 
select numorder, nomnom, sum(quant) as quant, min(date1), max(date1)
from itemProdShip
group by 	
	numorder, nomnom
;






if exists (select 1 from sysviews where viewname = 'itemBranShip' and vcreator = 'dba') then
	drop view itemBranShip;
end if;


create view itemBranShip (numorder, nomnom, quant, date1, scope)
as 
select r.numorder, r.nomnom, r.quant, r.date1, 'p'
from itemProdShip r
	union all
select r.numorder, r.nomnom, r.quant, r.date1, 'b'
from itemSellShip r
;



if exists (select 1 from sysviews where viewname = 'isumBranShip' and vcreator = 'dba') then
	drop view isumBranShip;
end if;

create view isumBranShip (numorder, nomnom, quant)

as 
-- объединение все вхождения номенклатуры в заказ через изделия или отгрузки в одну строку
-- имеем общее к-во по отгруженной номенклатуре всего заказа,
-- количество для нештучной ном-ры - в производственных единицах (дм)

select numorder, nomnom, sum(quant) as quant
from
	itemBranShip
group by
	numorder, nomnom
;


--=====================================================
--		Proc	сокращение от Processed:"незавершека"
--=====================================================

if exists (select 1 from sysviews where viewname = 'itemProdProc' and vcreator = 'dba') then
	drop view itemProdProc;
end if;

create view itemProdProc (numorder, nomnom, quant, statusid, firmname, ventureid, date1, ceh, manag)
as 
select 
	r.numdoc, r.nomnom, r.quant, o.statusid, f.name, o.ventureid, d.xdate, ceh, manag
from sdmc r
join sdocs d on r.numdoc = d.numdoc and r.numext = d.numext
join orders o on o.numorder = r.numdoc
join guidefirms f on f.firmid = o.firmid
join guidemanag m on m.managid = o.managid
join guideceh c on c.cehid = o.cehid
where o.statusid < 6
--left join itemProdShip s on s.numorder = r.numorder and s.nomnom = r.nomnom
;



if exists (select 1 from sysviews where viewname = 'itemSellProc' and vcreator = 'dba') then
	drop view itemSellProc;
end if;

create view itemSellProc (numorder, nomnom, quant, statusid, firmname, ventureid, date1, ceh, manag)
as 
select 
	r.numdoc, r.nomnom, r.quant, o.statusid, f.name, o.ventureid, d.xdate, 'Sell', m.manag
from sdmc r
join sdocs d on r.numdoc = d.numdoc and r.numext = d.numext
join bayorders o on o.numorder = r.numdoc
join guidemanag m on m.managid = o.managid
join bayguidefirms f on f.firmid = o.firmid
where statusid < 6
--left join itemProdShip s on s.numorder = r.numorder and s.nomnom = r.nomnom
;


if exists (select 1 from sysviews where viewname = 'itemBranProc' and vcreator = 'dba') then
	drop view itemBranProc;
end if;


create view itemBranProc (numorder, nomnom, quant, scope, statusid)
as 
select r.numorder, r.nomnom, r.quant, 'p', r.statusid
from itemProdProc r
		union all
select r.numorder, r.nomnom, r.quant, 'p', r.statusid
from itemSellProc r
;




if exists (select 1 from sysviews where viewname = 'isumProdProc' and vcreator = 'dba') then
	drop view isumProdProc;
end if;

create view isumProdProc (numorder, nomnom, quant, statusid, firmname, ventureid, date1, date2, ceh, manag)
as 
select 
	r.numorder, r.nomnom, sum(r.quant) as quant, r.statusid, r.firmname, r.ventureid, min(r.date1), max(date1), ceh, manag
from itemProdProc r
group by r.numorder, r.nomnom, r.statusid, r.firmname, r.ventureid, ceh, manag
having quant > 0
;


if exists (select 1 from sysviews where viewname = 'isumSellProc' and vcreator = 'dba') then
	drop view isumSellProc;
end if;

create view isumSellProc (numorder, nomnom, quant, statusid, firmname, ventureid, date1, date2, ceh, manag)
as 
select 
	r.numorder, r.nomnom, sum(r.quant) as quant, r.statusid, r.firmname, r.ventureid, min(r.date1), max(r.date1), ceh, manag
from itemSellProc r
group by r.numorder, r.nomnom, r.statusid, r.firmname, r.ventureid, ceh, manag
having quant > 0
;




if exists (select 1 from sysviews where viewname = 'isumBranProc' and vcreator = 'dba') then
	drop view isumBranProc;
end if;


create view isumBranProc (numorder, nomnom, quant, date1, date2, scope, statusid, firmname, ventureid, ceh, manag)
as
select r.numorder, r.nomnom, r.quant - isnull(s.quant, 0), r.date1, r.date2, 'p', r.statusid, r.firmname, r.ventureid, ceh, manag
from isumProdProc r
left join isumProdShip s on s.numorder = r.numorder and s.nomnom = r.nomnom
		union all
select r.numorder, r.nomnom, r.quant - isnull(s.quant, 0), r.date1, r.date2, 'b', r.statusid, r.firmname, r.ventureid, ceh, manag
from isumSellProc r
left join isumSellShip s on s.numorder = r.numorder and s.nomnom = r.nomnom
;


if exists (select 1 from sysviews where viewname = 'orderBranProc' and vcreator = 'dba') then
	drop view orderBranProc;
end if;


create view orderBranProc(numorder, sm_processed, statusid, date1, date2, firmname, scope, ceh, manag)
as
select r.numorder
	, sum(r.quant * n.cost / n.perlist) as quant
	, r.statusid , min(r.date1), max(r.date2), r.firmname, r.scope, ceh, manag
from isumBranProc r
join sguidenomenk n on n.nomnom = r.nomnom
group by r.numorder, r.statusid, r.firmname, r.scope, ceh, manag
having round(quant, 2) > 0
;




--===================================================================
--	Rsrv	сокращение от Reserved зарезервированная номенклатура
--===================================================================

if exists (select 1 from sysviews where viewname = 'itemProdRsrv' and vcreator = 'dba') then
	drop view itemProdRsrv;
end if;

create view itemProdRsrv (numorder, nomnom, quant, quant_rele)
as
select 
r.numdoc, r.nomnom, r.quantity, sum(isnull(d.quant, 0))
from sdmcrez r
left join sdmc d on d.numdoc = r.numdoc and d.nomnom = r.nomnom
join orders o on o.numorder = r.numdoc
group by r.numdoc, r.nomnom, r.quantity
;


if exists (select 1 from sysviews where viewname = 'isumProdRsrv' and vcreator = 'dba') then
	drop view isumProdRsrv;
end if;

create view isumProdRsrv (numorder, nomnom, quant, status, date1, manager, client, note, ceh, sm_zakazano, sm_paid)
as
select 
r.numorder, r.nomnom, r.quant - r.quant_rele as x_quant, s.status, o.indate
, m.manag, f.name, o.product, c.ceh, o.ordered, o.paid
from itemProdRsrv r
join orders o on o.numorder = r.numorder
join guidestatus s on s.statusid = o.statusid
left join guidemanag m on m.managid = o.managid
left join guidefirms f on f.firmid = o.firmid
left join guideceh c on c.cehid = o.cehid

where abs(round(x_quant, 2)) > 0.01
;


if exists (select 1 from sysviews where viewname = 'itemSellRsrv' and vcreator = 'dba') then
	drop view itemSellRsrv;
end if;

create view itemSellRsrv (numorder, nomnom, quant, quant_rele, date1)
as
select 
r.numdoc, r.nomnom, r.quantity, sum(isnull(d.quant, 0)), o.indate
from sdmcrez r
left join sdmc d on d.numdoc = r.numdoc and d.nomnom = r.nomnom
join bayorders o on o.numorder = r.numdoc 
group by r.numdoc, r.nomnom, r.quantity, o.indate
;

if exists (select 1 from sysviews where viewname = 'isumSellRsrv' and vcreator = 'dba') then
	drop view isumSellRsrv;
end if;

create view isumSellRsrv (numorder, nomnom, quant, status, date1, manager, client, note, ceh, sm_zakazano, sm_paid)
as
select
	r.numorder, r.nomnom, r.quant - r.quant_rele as x_quant, s.status, r.date1
	, m.manag, f.name, null, 'Продажа', o.ordered, o.paid
from itemSellRsrv r
join bayorders o on o.numorder = r.numorder
join guidestatus s on s.statusid = o.statusid
left join guidemanag m on m.managid = o.managid
left join bayguidefirms f on f.firmid = o.firmid
where abs(round(x_quant, 2)) > 0.01
;



if exists (select 1 from sysviews where viewname = 'itemFlawRsrv' and vcreator = 'dba') then
	drop view itemFlawRsrv;
end if;

create view itemFlawRsrv (numorder, nomnom, quant, date1, note)
as
select d.numdoc, r.nomnom, r.quantity, d.xdate, d.note
from sdocs d
join sdmcrez r on r.numdoc = d.numdoc
where d.numext = 0
;

if exists (select 1 from sysviews where viewname = 'isumFlawRsrv' and vcreator = 'dba') then
	drop view isumFlawRsrv;
end if;

create view isumFlawRsrv (numorder, nomnom, quant, status, date1, manager, client, note, ceh, sm_zakazano, sm_paid)
as
select
r.numorder, r.nomnom, r.quant, null, r.date1, null, 'Списание', null, r.note, null, null
from itemFlawRsrv r
;




if exists (select 1 from sysviews where viewname = 'isumBranRsrv' and vcreator = 'dba') then
	drop view isumBranRsrv;
end if;

create view isumBranRsrv (numorder, nomnom, quant, status, date1, manager, client, note, ceh, sm_zakazano, sm_paid, scope)
as
select 
	*, 'p'
from 
	isumProdRsrv
		union all
select 
	*, 'b'
from 
	isumSellRsrv
		union all
select 
	*, 'f'
from 
	isumFlawRsrv
;


if exists (select 1 from sysviews where viewname = 'orderBranRsrv' and vcreator = 'dba') then
	drop view orderBranRsrv;
end if;

create view orderBranRsrv (numorder, nomnom, quant, date1, manager, client, note, ceh, sm_zakazano, sm_paid, scope, status)
as
select 
	r.numorder, r.nomnom, (r.quant / n.perlist), r.date1, r.manager, r.client, r.note, r.ceh, r.sm_zakazano, r.sm_paid, r.scope, r.status
from isumBranRsrv r
join sguidenomenk n on n.nomnom = r.nomnom
;
/*


if exists (select 1 from sysviews where viewname = 'isumBranRsrv' and vcreator = 'dba') then
	drop view isumBranRsrv;
end if;

create view isumBranRsrv (numorder, nomnom, quant)
as
select 
o.numorder, o.nomnom, round((o.quant - isnull(s.quant, isnull(p.quant, 0))), 2) as r_quant
from isumBranOrde o
left join isumBranRele s on o.numorder = s.numorder and o.nomnom = s.nomnom
left join isumBranProc p on p.numorder = o.numorder and p.nomnom = o.nomnom
where r_quant <> 0 and o.statusid < 6
;

select round(n.cost / perlist *(o.quant - isnull(s.quant, isnull(p.quant, 0))), 2) as r_sum
, round((o.quant - isnull(s.quant, isnull(p.quant, 0))), 2) as r_quant
, o.*, s.quant as s, p.quant as p, n.perlist, n.cost, n.nomname
from isumBranOrde o
join sguidenomenk n on n.nomnom = o.nomnom
left join isumBranShip s on o.numorder = s.numorder and o.nomnom = s.nomnom
left join isumBranProc p on p.numorder = o.numorder and p.nomnom = o.nomnom
where r_quant <> 0 and o.statusid < 6
*/



-- Отчет А: строки дебитор/кредитор

if exists (select 1 from sysviews where viewname = 'vDebitorKreditor' and vcreator = 'dba') then
	drop view vDebitorKreditor;
end if;

create 
    view vDebitorKreditor (numorder, name, k, d, type)
as

SELECT numorder, f.name, (isnull(if isnull(paid, 0) > isnull(shipped, 0) then isnull(paid, 0) - isnull(shipped, 0) endif , 0)) AS k
, (isnull(if isnull(paid, 0) < isnull(shipped, 0) then isnull(shipped, 0) - isnull(paid, 0) endif , 0)) AS d, 't'
--, ordered, paid, shipped, statusid
from orders o
join guidefirms f on o.firmid = f.firmid
where (k != 0 or d != 0)
and statusid < 6
        union all
SELECT o.numorder, f.name, (isnull(if isnull(paid, 0) > isnull(cenaTotal, 0) then isnull(paid, 0) - isnull(cenaTotal, 0) endif , 0)) AS k
, (isnull(if isnull(paid, 0) < isnull(cenaTotal, 0) then isnull(cenaTotal, 0) - isnull(paid, 0) endif , 0)) AS d, 'b'
--, ordered, paid, cenaTotal
from bayorders o
join bayguidefirms f on o.firmid = f.firmid
left join orderSellShip ot on ot.numorder = o.numorder
where (k != 0 or d != 0)
and o.statusid < 6
;





--=====================================================
--		Тяжелое наследство (еще от Димы)
--=====================================================


if exists (select 1 from sysviews where viewname = 'wCloseNomenk' and vcreator = 'dba') then
	drop view wCloseNomenk;
end if;


create VIEW wCloseNomenk (
	numDoc,
	nomNom,
	quantity,
	Sum_quant
)	as
SELECT 
	r.numDoc, 
	r.nomNom, 
	r.quantity, 
	Sum(isnull(d.quant, 0)) AS Sum_quant
FROM 
	sDMCrez r
	LEFT JOIN sDMC d ON r.nomNom = d.nomNom AND r.numDoc = d.numDoc
GROUP BY r.numDoc, r.nomNom, r.quantity;



if exists (select 1 from sysviews where viewname = 'wCloseNomenk2' and vcreator = 'dba') then
	drop view wCloseNomenk2;
end if;

create VIEW wCloseNomenk2 (
	numDoc,
	nomNom,
	quantity,
	Sum_quant
)	as
SELECT r.numDoc, r.nomNom, r.quantity, Sum(isnull(d.quant, 0)) AS Sum_quant
FROM 
	sDMCrez r 
	INNER JOIN Orders o ON r.numDoc = o.numOrder
	LEFT JOIN sdmc d ON r.nomNom = d.nomNom AND r.numDoc = d.numDoc
GROUP BY r.numDoc, r.nomNom, r.quantity;



if exists (select 1 from sysviews where viewname = 'wCloseNomenk3' and vcreator = 'dba') then
	drop view wCloseNomenk3;
end if;

create VIEW wCloseNomenk3 (
	numDoc,
	nomNom,
	quantity,
	Sum_quant
)	as
SELECT 
	r.numDoc, 
	r.nomNom, 
	r.quantity, 
	Sum(isnull(d.quant, 0)) AS Sum_quant
FROM 
	sDMCrez r 
	INNER JOIN Orders o ON r.numDoc = o.numOrder
	LEFT JOIN sdmc d ON r.nomNom = d.nomNom AND r.numDoc = d.numDoc
GROUP BY r.numDoc, r.nomNom, r.quantity, r.curQuant
HAVING r.curQuant > 0;
--ORDER BY r.numDoc;


if exists (select 1 from sysviews where viewname = 'mm_orders' and vcreator = 'dba') then
	drop view mm_orders;
end if;


create view mm_orders as
select numorder, indate, invoice, id_jscet, ventureid 
from orders o
where (((char_length(o.invoice) = 6 and o.invoice != 'счет ?') or char_length(invoice) = 5) and substring(invoice, 2, 1) = '5') or ventureid = 2
    union
select numorder, indate, invoice, id_jscet, ventureid 
from bayorders o
where (((char_length(o.invoice) = 6 and o.invoice != 'счет ?') or char_length(invoice) = 5) and substring(invoice, 2, 1) = '5') or ventureid = 2 
;

if exists (select 1 from sysviews where viewname = 'pm_orders' and vcreator = 'dba') then
	drop view pm_orders;
end if;


create view pm_orders as
select numorder, indate, invoice, id_jscet, ventureid 
from orders o
where (((char_length(o.invoice) = 6 and o.invoice != 'счет ?') or char_length(invoice) = 5) and substring(invoice, 2, 1) = '0') or ventureid = 1
    union
select numorder, indate, invoice, id_jscet, ventureid 
from bayorders o
where (((char_length(o.invoice) = 6 and o.invoice != 'счет ?') or char_length(invoice) = 5) and substring(invoice, 2, 1) = '0') or ventureid = 1  
;

--------------------------------- единая вью по всем заказам включая продажи ---------------------

if exists (select 1 from sysviews where viewname = 'all_orders' and vcreator = 'dba') then
	drop view all_orders;
end if;

create view all_orders (numorder, tp, xdate, statusid, firmName) 
-- заказы производства и продаж единым списком
as 
select numorder, 'orders', indate, statusid, f.name
from orders o
join guidefirms f on f.firmid = o.firmid
	union 
select numorder, 'bayorders', indate, statusid, f.name
from bayorders o
join bayguidefirms f on f.firmid = o.firmid
;


if exists (select 1 from sysviews where viewname = 'wEbSvodka' and vcreator = 'dba') then
	drop view wEbSvodka;
end if;

create VIEW wEbSvodka (
	xLogin, 
	numOrder, 
	StatusId, 
	outDateTime, 
	Problem, 
	Logo, 
	Product, 
	ordered, 
	paid, 
	shipped, 
	Name, 
	Manag, 
	DateRS
	
) as
SELECT 
	GuideFirms.xLogin, 
	Orders.numOrder, 
	Orders.StatusId, 
	Orders.outDateTime, 
	GuideProblem.Problem, 
	Orders.Logo, 
	Orders.Product, 
	Orders.ordered, 
	Orders.paid, 
	Orders.shipped, 
	GuideFirms.Name, 
	GuideManag.Manag, 
	Orders.DateRS
FROM 
	GuideStatus 
	INNER JOIN (GuideProblem 
		INNER JOIN (GuideFirms 
			INNER JOIN (GuideManag 
				INNER JOIN Orders ON GuideManag.ManagId = Orders.ManagId) 
			ON GuideFirms.FirmId = Orders.FirmId) 
		ON GuideProblem.ProblemId = Orders.ProblemId) 
	ON GuideStatus.StatusId = Orders.StatusId
WHERE Orders.StatusId not in (6, 7)
;
