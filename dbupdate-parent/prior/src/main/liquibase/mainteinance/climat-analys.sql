begin
	-- 
	create variable @curId integer;
	create variable @resultHeaderId integer;
	create variable @analysTemplateId integer;
	create variable @ByColumnId integer;
	create variable @ByRowId integer;
	create variable @ParamFrameId integer;
	create variable @itemTypeId integer;
	create variable @resultTitleParamId integer;




	insert into nAnalysBootingParam(name, description) values ('parametersFrame', 'Имя фрейма с параметрами в окне диалога.');
	set @ParamFrameId = @@identity;

	insert into nAnalysBootingParam(name, description) values ('resultTitle', 'Строка форматирования заголовка результата. Пример: Материалы купленные фирмой ${clientId} за период с ${startDate} по ${endDate}');
	set @resultTitleParamId = @@identity;


	-- frame с параметрами для всех прочих отчетов
	insert into nAnalysBooting (templateId, paramValue, paramId)
	select t.id, 'default', @ParamFrameId
	from nAnalysTemplate t;


	insert into nResultHeader (name) select 'climat';
	set @resultHeaderId = @@identity;


	
	insert into nResultColumns (headerId, columnId, sort, hidden)
	select @resultHeaderId, rcd.id, 0, 1
	from nResultColumnDef rcd where rcd.name = 'nomnom';

	insert into nResultColumns (headerId, columnId, sort, hidden)
	select @resultHeaderId, rcd.id, 10, 0
	from nResultColumnDef rcd where rcd.name = 'nomnom';

	insert into nResultColumns (headerId, columnId, sort, hidden)
	select @resultHeaderId, rcd.id, 20, 0
	from nResultColumnDef rcd where rcd.name = 'nomname';

	insert into nResultColumns (headerId, columnId, sort, hidden)
	select @resultHeaderId, rcd.id, 30, 0
	from nResultColumnDef rcd where rcd.name = 'edizm';

	insert into nResultColumns (headerId, columnId, sort, hidden)
	select @resultHeaderId, rcd.id, 35, 0
	from nResultColumnDef rcd where rcd.name = 'cena';

	insert into nResultColumns (headerId, columnId, sort, hidden)
	select @resultHeaderId, rcd.id, 40, 0
	from nResultColumnDef rcd where rcd.name = 'materialQty';

	insert into nResultColumns (headerId, columnId, sort, hidden)
	select @resultHeaderId, rcd.id, 50, 0
	from nResultColumnDef rcd where rcd.name = 'materialSaled';

	insert into nItemType (itemType) values ('client');
	set @itemTypeId = @@identity;

--	insert into nParamType (paramType, paramClass, paramKey, itemTypeId) 
--	select 'clientId', 'id', null, @itemTypeId;

	insert into nAnalysTemplate (sqlFunction, headerId, sqlHeader) values ('n_list_climat_by_periods', @resultHeaderId, 'call n_fill_periods(v_begin, v_end, v_sub_token)');
	set @analysTemplateId = @@identity;

	insert into nAnalysBooting (templateId, paramValue, paramId)
	select @analysTemplateId, 1, p.id
	from nAnalysBootingParam p 
	where p.name = 'noRowDetail';

	insert into nAnalysBooting (templateId, paramValue, paramId)
	select @analysTemplateId, 'поз. номенклатуры', p.id
	from nAnalysBootingParam p 
	where p.name = 'totalQtyLabel';

	

	insert into nAnalysBooting (templateId, paramValue, paramId)
	select @analysTemplateId, 'Материалы купленные фирмой "${clientName}"', @resultTitleParamId;

	insert into nAnalysCategory(name, name_ru, parentId, byrow_flag, bycolumn_flag) values ('climat','Все материалы клиента',null, 1, 0);
	set @ByRowId = @@identity;

	insert into nAnalys (byrow, bycolumn, templateId, application)
	select @ByRowId, ac.id, @analysTemplateId, 'bay'
	from nAnalysCategory ac
	where ac.parentId = 1;

	-- frame с параметрами для нового параметра
	insert into nAnalysBooting (templateId, paramValue, paramId)
	select @analysTemplateId, 'climat', @ParamFrameId;

	insert into nAnalysBooting (templateId, paramValue, paramId)
	select @analysTemplateId, 'nomnom', p.id
	from nAnalysBootingParam p
	where 
		p.name = 'groupSelectorColumn';


	insert into nAnalysBooting (templateId, paramValue, paramId)
	select @analysTemplateId, 'periodId', p.id
	from nAnalysBootingParam p
	where 
		p.name = 'periodId4detail';

end;
