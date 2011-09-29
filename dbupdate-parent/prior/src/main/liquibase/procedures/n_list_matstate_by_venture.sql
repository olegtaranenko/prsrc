ALTER PROCEDURE "DBA"."n_list_matstate_by_venture" (
	  p_begin         date
	, p_end           date
	, p_period_type   varchar(20) -- p_sub_token
	, p_rowId         integer
	, p_columnId      integer
)
begin
	
	declare c_begin date;

	create table #saldo(nomnom varchar(20), id integer, debit float, kredit float, periodId integer, xDate date null);
	create table #itogo(nomnom varchar(20), id integer, debit float, kredit float, periodId integer);
	create table #norm_itogo(nomnom varchar(20), id integer, debit float, kredit float, periodId integer);

	create table #turn_periods (
		  periodId      integer      not null
		, st            date         not null
		, en            date         not null
	);




	--TODO дать возможность из этой процедуры доступ к установка фильтра.
--	set c_begin = convert(date, n_get_booting_param(p_filterId, 'minDate'));
	set c_begin = '20000101';

	message 'begin = ', p_begin to client;
	message 'end = ', p_end to client;

	if p_begin <> c_begin then
		insert into #turn_periods (periodId, st, en) select 0, c_begin, p_begin;
	end if;

	insert into #turn_periods (periodId, st, en) select 1, p_begin, p_end;

	
	insert into #saldo (nomnom, id, debit, kredit, periodId)
	select r_nomnom, r_ventureid, sum(r_debit) as debit, sum(r_kredit) as kredit, r_periodid
	from dummy
		join (
			select
				 m.nomnom as r_nomnom
				, if n.destid <= -1001 then 
					quant
				else
					0
				endif
	    			as r_debit
				, if n.destid <= -1001 then 
					0
				else
					quant
				endif
	    			as r_kredit
    			, n.ventureid as r_ventureid
				, n.xdate as r_input_date
				, tp.periodId as r_periodId
        	from sdocs n
    		join sdmc m on n.numdoc = m.numdoc and n.numext = m.numext 
		join sguidesource s on s.sourceId = n.sourId
    		join sguidesource d on d.sourceId = n.destId
    		join system sys on 1 = 1
    		join guideventure v on v.id_analytic = sys.id_analytic_default
    		left join orders o on o.numorder = n.numdoc
    		left join bayorders bo on bo.numorder = n.numdoc
			join #turn_periods tp on n.xdate >= tp.st and n.xdate < tp.en
			where
    			convert(date, n.xDate) <= isnull(p_end, convert(date, n.xDate))
			and (n.sourid > -1001 or n.destid > -1001)
    	) x on 1=1
	group by r_nomnom, r_ventureid, r_periodid;


	insert into #saldo (nomnom, id, debit, kredit, periodid)
    select m.nomnom, srcVentureId, 0, sum(m.quant) as kredit, tp.periodid
			from sdmcventure m
			join sdocsventure n on m.sdv_id = n.id and n.cumulative_id is not null
			join #turn_periods tp on n.ndate >= tp.st and n.ndate < tp.en
			where n.nDate < p_end
			group by 
				m.nomnom, srcVentureId, tp.periodid;

	insert into #saldo (nomnom, id, debit, kredit, periodid)
    select m.nomnom, dstVentureId, sum(m.quant) as kredit, 0, tp.periodid
			from sdmcventure m
			join sdocsventure n on m.sdv_id = n.id and n.cumulative_id is not null
			join #turn_periods tp on n.ndate >= tp.st and n.ndate < tp.en
			where n.nDate < p_end
			group by 
				m.nomnom, dstVentureId, tp.periodid;


	insert into #itogo (nomnom, id, debit, kredit, periodid)
	select s.nomnom, id, sum(debit / n.perlist), sum(kredit / n.perlist), s.periodid 
	from #saldo s
	join sguidenomenk n on n.nomnom = s.nomnom
	group by 
		s.nomnom, s.id, s.periodid;

	insert into #results (nomnom, periodId, nomname, edizm, cena)
	select distinct i.nomnom, i.id, trim(n.cod + ' ' + nomname + ' ' + n.size), n.ed_izmer2, n.cost
	from #itogo i
	join sguidenomenk n on n.nomnom = i.nomnom 
	order by i.nomnom, i.id
	;

	update #results set matInQty = isnull(i.debit, 0) - isnull(i.kredit, 0)
		, matOutQty = isnull(i.debit, 0) - isnull(i.kredit, 0)
		, sumOut = (isnull(i.debit, 0) - isnull(i.kredit, 0)) * cena
	from #itogo i
	where #results.nomnom = i.nomnom and #results.periodid = i.id and i.periodid = 0;


	update #results set 
		  matInTurn = isnull(i.debit, 0)
		, matOutTurn = isnull(i.kredit, 0)
		, matOutQty = isnull(matInQty, 0) + isnull(i.debit, 0) - isnull(i.kredit, 0)
		, sumOut = (isnull(matInQty, 0) + isnull(i.debit, 0) - isnull(i.kredit, 0)) * cena
	from #itogo i
	where #results.nomnom = i.nomnom and #results.periodid = i.id and i.periodid = 1;

end