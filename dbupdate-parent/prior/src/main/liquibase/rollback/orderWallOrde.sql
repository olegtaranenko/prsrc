if exists (select 1 from sysviews where viewname = 'orderWallOrde' and vcreator = 'dba') then
	drop view orderWallOrde;
end if;


create view orderWallOrde (
	  outdate
	, numorder
	, type
	, code
	, nomnom
	, prId
	, prExt
	, cenaEd
	, quant
	, name
	, ventureId
	, statusid
	, werkid
	, itemName
	, edIzm
) as
select 
	 o.outdatetime
	,o.numorder
	,'p'
	,wf_make_prnm(p.prName, pi.prext)
	,p.prName
	,pi.prId
	,pi.prExt
	,cenaed
	,quant 
	,f.name
	,o.ventureId
	,o.statusid
	,o.werkId
	,wf_make_invnm(p.prDescript, p.prSize)
	, 'רע.'
from
	orders o
join xpredmetybyizdelia pi on pi.numorder = o.numorder
left join firmguide f on f.firmid = o.firmid
join sguideproducts p on p.prId = pi.prId
	union all

select 
	 o.outdatetime
	,o.numorder
	,'n'
	,pn.nomnom
	,pn.nomnom
	,null
	,null
	,cenaed * n.perlist
	,quant  / n.perlist
	,f.name
	,o.ventureId
	,o.statusid
	,o.werkId
	,wf_make_invnm(n.nomName, n.size, n.cod)
	,ed_izmer2
from
	orders o
join xpredmetybynomenk pn on pn.numorder = o.numorder
join sguidenomenk n on n.nomnom = pn.nomnom
left join firmguide f on f.firmid = o.firmid;