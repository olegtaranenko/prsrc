begin
	create variable @resultHeaderId integer;
	create variable @curId integer;
	create variable @analysTemplateId integer;

	insert into nResultHeader (name) select 'nomnom';
	set @resultHeaderId = @@identity;
	
	insert into nAnalysTemplate (sqlFunction, headerId, sqlHeader) values ('n_list_nomnom_by_firm', @resultHeaderId, 'call n_fill_firms(p_filterId, v_begin, v_end)');
	set @analysTemplateId = @@identity;

	insert into nAnalysCategory(name, name_ru, parentId, byrow_flag, bycolumn_flag) values ('nomnom','Материалы',null, 1, 0);
	set @curId = @@identity;

	insert into nAnalys (byrow, bycolumn, templateId, application) 
	select @curId, c.id, @analysTemplateId, 'bay'
	from nAnalysCategory c where c.name = 'firm';

	update nAnalysCategory set bycolumn_flag = 1 
	where name = 'firm';

	insert into nAnalysBooting (templateId, paramValue, paramId)
	select @analysTemplateId, 'nomnom', p.id
	from nAnalysBootingParam p
	where 
		p.name = 'groupSelectorColumn';

	insert into nAnalysBooting (templateId, paramValue, paramId)
	select @analysTemplateId, 'firmId', p.id
	from nAnalysBootingParam p
	where 
		p.name = 'periodId4detail';

	-- пока детализация для отчета невозможна
	insert into nAnalysBooting (templateId, paramValue, paramId)
	select @analysTemplateId, '1', p.id
	from nAnalysBootingParam p
	where 
		p.name = 'noRowDetail';


	insert into nAnalysBooting (templateId, paramValue, paramId)
	select @analysTemplateId, 'номенклатур(а/ы)', p.id
	from nAnalysBootingParam p
	where 
		p.name = 'totalQtyLabel';
		
	insert into nAnalysBooting (templateId, paramValue, paramId)
	select @analysTemplateId, '20000101', p.id
	from nAnalysBootingParam p
	where 
		p.name = 'minDate';

	insert into nAnalysBooting (templateId, paramValue, paramId)
	select @analysTemplateId, '30000101', p.id
	from nAnalysBootingParam p
	where 
		p.name = 'maxDate';


	insert into nResultColumns (headerId, columnId, sort, hidden)
	select @resultHeaderId, rcd.id, 10, 0
	from nResultColumnDef rcd
	where rcd.name = 'nomnom';

	insert into nResultColumns (headerId, columnId, sort, hidden)
	select @resultHeaderId, rcd.id, 100, 0
	from nResultColumnDef rcd
	where rcd.name = 'nomname';

	insert into nResultColumns (headerId, columnId, sort, hidden)
	select @resultHeaderId, rcd.id, 200, 0
	from nResultColumnDef rcd
	where rcd.name = 'edizm';

	insert into nResultColumns (headerId, columnId, sort, hidden)
	select @resultHeaderId, rcd.id, 300, 0
	from nResultColumnDef rcd
	where rcd.name = 'cena';

	insert into nResultColumns (headerId, columnId, sort, hidden)
	select @resultHeaderId, rcd.id, 400, 0
	from nResultColumnDef rcd
	where rcd.name = 'matOutQty';


end;
