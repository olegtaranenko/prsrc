if exists (select 1 from sysprocedure where proc_name = 'n_production_visitors_rebuild') then
	drop procedure n_production_visitors_rebuild;
end if;


CREATE PROCEDURE n_production_visitors_rebuild (
--	  p_year         date
)
begin


	declare v_current_year  int;
	declare v_start_year  int;
	declare v_start_year_str    varchar(8);
	declare v_end_year_str    varchar(8);
	

	select year(now()) + 1 into v_current_year ;
	select v_current_year - 4 into v_start_year;
	
	set v_start_year_str = convert(varchar(4), v_start_year) + '0101';
	set v_end_year_str = convert(varchar(4), v_current_year) + '0101';


	create table #periods (
		  periodId      integer      default autoincrement
		, klassId       integer      null
		, ventureId     integer      null
		, label         varchar(32)  null
		, st            date         null
		, en            date         null
		, year          integer      null
	);

	call n_fill_periods(v_start_year_str,  v_end_year_str, 'year');
	
	
	update #periods set st = '20000101' where year = v_start_year;
--	delete from #periods where year < v_current_year;

--	select * from #periods;

	create table #orders(
		  numorder   integer primary key
		, indate     date
		, periodid   integer	null
		, firmId     integer
	);


	insert into #orders (
		numorder, 
		indate, 
		firmId
	)
	select 
		o.numorder, 
		o.indate, 
		o.firmId
	from 
		orders o
	where werkId = 2;

--	select count(*) from #orders;

	update #orders s set s.periodId = p.periodId
	from #periods p 
	where 
		s.indate >= p.st and s.inDate < p.en
	;

-- 	select * from #orders;

	create table #summary (
		visits integer 
		, periodId integer null
		, firmId integer
	);

	
	insert into #summary (visits, periodId, firmId)
	select count(*), periodId, firmId 
	from #orders
	group by periodid, firmid;

-- 	select * from #summary;
	
--	update FirmGuide set year01 = 0, year02 = 0, year03 = 0, year04 = 0;

	update FirmGuide 
	set year01 = visits
	from #summary
	where FirmGuide.firmId = #summary.firmId and #summary.periodId = 1;
	message 'rowcount = ', @@rowcount to client;

	
	update FirmGuide set year02 = visits
	from #summary
	where FirmGuide.firmId = #summary.firmId and #summary.periodId = 2;
	
	update FirmGuide set year03 = visits
	from #summary
	where FirmGuide.firmId = #summary.firmId and #summary.periodId = 3;
	
	update FirmGuide set year04 = visits
	from #summary
	where FirmGuide.firmId = #summary.firmId and #summary.periodId = 4;


commit;
end;
