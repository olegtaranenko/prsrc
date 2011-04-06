if exists (select 1 from sysviews where viewname = 'itemWallOrde' and vcreator = 'dba') then
	drop view itemWallOrde;
end if;


create view itemWallOrde (
	  numorder
	, type
	, code
	, nomnom
	, prId
	, prExt
	, cenaEd
	, quant
	, werkid
	, itemName
	, edIzm
	, edIzmList
	, firmId
) as
select 
	 o.numorder
	,'p'
	,wf_make_prnm(p.prName, pi.prext)
	,p.prName
	,pi.prId
	,pi.prExt
	,cenaed
	,quant 
	,o.werkId
	,wf_make_invnm(p.prDescript, p.prSize)
	, 'רע.'
	, 'רע.'
	,o.FirmId
from
	orders o
join xpredmetybyizdelia pi on pi.numorder = o.numorder
join sguideproducts p on p.prId = pi.prId

	union all

select 
	 o.numorder
	,'n'
	,pn.nomnom
	,pn.nomnom
	,null
	,null
	,cenaed * n.perlist
	,quant  / n.perlist
	,o.werkId
	,wf_make_invnm(n.nomName, n.size, n.cod)
	,ed_izmer
	,ed_izmer2
	,o.FirmId
from
	orders o
join xpredmetybynomenk pn on pn.numorder = o.numorder
join sguidenomenk n on n.nomnom = pn.nomnom;