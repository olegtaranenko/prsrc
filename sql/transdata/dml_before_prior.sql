-- ��������� �� 14 �������. 
-- ����������� ������ ������ �������������
-- ������� ��������� ������ ����� ��������� ������ � ��������


--call bootstrap_blocking;

--delete from sdocsventure;

-- � ������� => ���
if not exists(select 1 from sys.syscolumns where creator = 'dba' and tname = 'sdocs' and cname = 'ventureId') then
	alter table sdocs add ventureId integer null;
	alter table sdocs add constraint ventureId foreign key (ventureId) references guideVenture (ventureId) on update cascade on delete set null;
end if;                                          

if exists (select 1 from systable where table_name = 'sdocsincome') then
	update sdocs set ventureid = 1 
	where numext = 255 and destid = -1001
	and sourId not in (34, 0);

	update sdocs d set ventureId = i.ventureId
	from sdocsIncome i
	where i.numdoc = d.numdoc and i.numext = d.numext;

	drop table sdocsIncome;
end if;

commit;
