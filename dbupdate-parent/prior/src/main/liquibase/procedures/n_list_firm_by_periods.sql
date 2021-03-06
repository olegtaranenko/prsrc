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

	if v_detail_fine = 1 then 
		select st, en 
		into v_begin, v_end
		from #periods where periodId = p_columnId;
	else
		set v_begin = p_begin;
		set v_end   = p_end;
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
		,prId        integer  null
		,prExt       integer  null
		,materialQty float    null
		,cenaEd      float    null
		,inDate      date
		,firmId      integer
		,klassid     integer
		,periodid    integer  null
		,priceToDate float    null
		,quantEd     float    null
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
    
		call n_sell_item_cost(v_begin, v_end, 1);

		update #orders s set s.hasMaterial = 1
		where 
			exists (select 1 from #sale_item i where i.numorder = s.numorder)
		;

		insert into #firm_has_mat
		select distinct firmId 
		from #orders s
		where s.hasMaterial = 1
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
			, o.orderQty            -- ����� ������� �� ������
			, o.orderPaid           -- ����� ����� ������� (��)
			, o.orderOrdered        -- ����� ����� �� �������
			, o.orderMatQty         -- �-�� �������, � ������� �������������� ��������� ���������
			, i.materialSaled    	-- ����� �� ��������� ���������� (���� ��� ���������� - ��������� ����� �����������)
			, f.name                -- �����
			, r.region
			, r.regionid
			, p.periodid
			, o.firmId
			, f.tools as oborud
		from #periods p 
		join (
			select sum(isnull(orderPaid, 0)) as orderPaid
				, count(*) as orderQty
				, firmId
				, periodId
				, sum(orderOrdered) as orderOrdered
				, sum(hasMaterial) as orderMatQty
			from 
				#orders
			group by 
				firmId, periodId
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
				o.firmId = i.firmId 
			and o.periodId = i.periodId
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


