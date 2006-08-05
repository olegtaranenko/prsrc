-- Изменения от 14 февраля. 
-- Перестроена логика работы взаимозачетов
-- Сводные накладные теперь будут храниться вместе с дневными


--call bootstrap_blocking;

--delete from sdocsventure;

-- в продакш => нет
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


-- в продакш => нет
if not exists(select 1 from sys.sysviews where viewname = 'all_orders') then
	create view all_orders (numorder, tp, xdate) as 
	select numorder, 'orders', indate from orders
		union 
	select numorder, 'bayorders', indate from bayorders;
end if;

commit;
