begin
	-- 
	create variable @curId integer;
	create variable @resultHeaderId integer;
	create variable @analysTemplateId integer;
	
	insert into nAnalysBootingParam(name, description) values ('firstVisit', '��������� �� ��� ������ ������ ������ �(���) ��������� ��������� �����');
	-- �������� ��� ���� ������� �������
	insert into nAnalysBooting (templateId, paramValue, paramId)
	select t.id, '1', @@identity
	from nAnalysTemplate t;

	insert into nAnalysBootingParam(name, description) values ('noRowDetail', '����������, ����� �� ������� �������� ������ ���� ��������������. ���� 1 �� ����������� ���, ������� �� ��������� ����������� ����������� ���������������.');
	insert into nAnalysBootingParam(name, description) values ('minDate', '������������ ��� ������ ��������� ���� ��� ���������� ���������');
	insert into nAnalysBootingParam(name, description) values ('maxDate', '������������ ��� ������ �������� ���� ��� ���������� ���������');


	insert into nResultHeader (name) select 'matstate';
	set @resultHeaderId = @@identity;
	
	insert into nAnalysTemplate (sqlFunction, headerId, sqlHeader) values ('n_list_matstate_by_venture', @resultHeaderId, 'call n_fill_ventures(p_filterId, v_begin, v_end)');
	set @analysTemplateId = @@identity;

	insert into nAnalysCategory(name, name_ru, parentId, byrow_flag, bycolumn_flag) values ('matstate','��������� ������',null, 1, 0);
	set @curId = @@identity;
	insert into nAnalysCategory(name, name_ru, parentId, byrow_flag, bycolumn_flag) values ('venture','�����������' ,null, 1, 1);
	insert into nAnalys (byrow, bycolumn, templateId) select @curId, @@identity, @analysTemplateId;
	
	
	insert into nAnalysBooting (templateId, paramValue, paramId)
	select @analysTemplateId, 'nomnom', p.id
	from nAnalysBootingParam p
	where 
		p.name = 'groupSelectorColumn';

	insert into nAnalysBooting (templateId, paramValue, paramId)
	select @analysTemplateId, 'ventureId', p.id
	from nAnalysBootingParam p
	where 
		p.name = 'periodId4detail';

	-- ���� ����������� ��� ������ ����������
	insert into nAnalysBooting (templateId, paramValue, paramId)
	select @analysTemplateId, '1', p.id
	from nAnalysBootingParam p
	where 
		p.name = 'noRowDetail';


	insert into nAnalysBooting (templateId, paramValue, paramId)
	select @analysTemplateId, '�������(�/�)', p.id
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

	insert into nResultColumnDef (name, name_ru, align, sort, hidden, headType)
	select                      'nomnom', '������.', '<', 0, 0, 1;
	insert into nResultColumns (headerId, columnId, sort, hidden)
	select @resultHeaderId, @@identity, 10, 0;
	
	insert into nResultColumnDef (name, name_ru, align, sort, hidden, headType)
	select                      'nomname', '�������� ������������', '<', 1, 0, 1;
	insert into nResultColumns (headerId, columnId, sort, hidden)
	select @resultHeaderId, @@identity, 100, 0;
	
	insert into nResultColumnDef (name, name_ru, align, sort, hidden, headType)
	select                      'edizm', '��.���.', '<', 1, 0, 1;
	insert into nResultColumns (headerId, columnId, sort, hidden)
	select @resultHeaderId, @@identity, 150, 0;
	
	insert into nResultColumnDef (name, name_ru, align, sort, hidden, headType, format)
	select                      'cena', '����.', '>', 1, 0, 1, '# ##0.00';
	insert into nResultColumns (headerId, columnId, sort, hidden)
	select @resultHeaderId, @@identity, 170, 0;
	
	insert into nResultColumnDef (name, name_ru, align, sort, hidden, headType, format)
	select                      'matInQty', '�� ������', '>', 1, 0, 2, '# ##0.00';
	insert into nResultColumns (headerId, columnId, sort, hidden)
	select @resultHeaderId, @@identity, 200, 0;
	
	insert into nResultColumnDef (name, name_ru, align, sort, hidden, headType, format)
	select                      'matInTurn', '������', '>', 2, 0, 2, '# ##0.00';
	insert into nResultColumns (headerId, columnId, sort, hidden)
	select @resultHeaderId, @@identity, 300, 0;
	
	insert into nResultColumnDef (name, name_ru, align, sort, hidden, headType, format)
	select                      'matOutTurn', '������', '>', 2, 0, 2, '# ##0.00';
	insert into nResultColumns (headerId, columnId, sort, hidden)
	select @resultHeaderId, @@identity, 350, 0;
	
	insert into nResultColumnDef (name, name_ru, align, sort, hidden, headType, format)
	select                      'matOutQty', '�� �����', '>', 3, 0, 2, '# ##0.00';
	insert into nResultColumns (headerId, columnId, sort, hidden)
	select @resultHeaderId, @@identity, 400, 0;

	insert into nResultColumnDef (name, name_ru, align, sort, hidden, headType, format)
	select                      'sumOut', '�����', '>', 3, 0, 2, '# ##0.00';
	insert into nResultColumns (headerId, columnId, sort, hidden)
	select @resultHeaderId, @@identity, 500, 0;

	commit;
end;
