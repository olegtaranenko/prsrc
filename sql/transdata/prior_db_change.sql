/**************************
* Изменить текущие таблицы схемы prr для связывания 
* значений с сущностями-аналогами из Comtec 
*/

if not exists (select 1 from systable where table_name = 'size') then
alter table BayGuideFirms add id_voc_names integer;
--alter table BayGuideProblem
alter table bayNomenkOut add (id_mat integer, id_jmat integer);
alter table BayOrders add (id_jscet integer, ventureid integer);
alter table BayRegion add id_voc_names integer;
alter table GuideCeh add id_voc_names integer;
alter table GuideFirms add id_voc_names integer;
alter table GuideManag add id_voc_names integer;
--alter table GuideProblem
--alter table GuideStatus
--alter table GuideTema
--alter table Itogi_CO2
--alter table Itogi_SUB
--alter table Itogi_YAG
alter table Orders add (id_jscet integer, lastModified datetime);
alter table OrdersInCeh add id_jscet integer;
alter table OrdersMO add id_jscet integer;
--alter table ReplaceRS
--alter table ResursCO2       	
--alter table ResursSUB
--alter table ResursYAG
alter table sDMC add id_mat integer;
alter table sDMCmov add id_mat integer;
alter table sDMCrez add (id_mat integer, id_scet integer);
alter table sDocs add id_jmat integer;
--alter table sGuideFormuls
alter table sGuideKlass add id_inv integer;
alter table sGuideNomenk add id_inv integer;
alter table sGuideProducts add id_inv integer;
alter table sGuideSeries add id_inv integer;
alter table sGuideSource add id_voc_names integer;
alter table sProducts add id_compl integer;

alter table system add (id_cur_rate integer, id_cur integer, trans_date datetime);

alter table xEtapByIzdelia add id_scet integer;
alter table xEtapByNomenk add id_scet integer;
alter table xPredmetyByIzdelia add (id_scet integer, id_inv integer);
alter table xPredmetyByIzdeliaOut add (id_mat integer, id_jmat integer);
alter table xPredmetyByNomenk add id_scet integer;
alter table xPredmetyByNomenkOut add (id_mat integer, id_jmat integer);
alter table xUslugOut add (id_mat integer, id_jmat integer);
alter table xVariantNomenc add id_scet integer;
alter table yBook add (id_xoz integer, ventureid integer);
--alter table yDebKreditor add (id_voc_names integer, ventureid integer);
--alter table yGuideDet
--alter table yGuideDetail
--alter table yGuidePurp
--alter table yGuidePurpose
/* 
Может и не нужно, если подсчета не повторяются для разных фирм
Нужно уточнить у бухгалтеров

alter table yGuideSchets add ventureid integer;
update yGuideSchets set ventureid = 0;
alter table yGuideSchets modify ventureid integer not null default 0;
*/

------------------------------------------
-- мусор от которого слудует избавиться --
--select * from sysforeignkeys where foreign_tname = 'sdocs'
/*
alter table sproducts drop foreign key sproducts_856;
alter table sproducts delete prid;

alter table sguideproducts drop foreign key sguideproducts_857;
alter table sguideproducts delete seriaid;
alter table sguideproducts drop foreign key sGuideFormulssGuideProducts;
alter table sguideproducts delete nomer;

alter table sguidenomenk drop foreign key sguideformulssguidenomenk;
alter table sguidenomenk drop foreign key sguideformulssguidenomenk1;
alter table sguidenomenk delete nomer;

alter table sdocs drop foreign key sguidesourcesdocs;
alter table sdocs drop foreign key sguidesourcesdocs1;
alter table sdocs delete sourceid;

alter table xvariantnomenc drop foreign key sProductsxVariantNomenc;
alter table xvariantnomenc delete productid;

alter table sguideklass drop foreign key sguideklass_851;
alter table sguideseries drop foreign key sguideseries_858;
*/

alter table ybook drop foreign key yBook_381;
alter table ybook drop foreign key yBook_382;
alter table ybook drop foreign key yBook_384;
alter table ybook drop foreign key yBook_385;
alter table yGuideDetail drop foreign key yGuideDetyGuideDetail;
alter table yGuideDetail drop foreign key yGuideDetail_383;
alter table yGuidePurpose drop foreign key yGuidePurpose_386;
alter table yGuidePurpose drop foreign key yGuidePurpose_387;
alter table yGuidePurpose drop foreign key yGuidePurpyGuidePurpose;

alter table ybook delete pId;
alter table ybook delete id;
alter table ybook delete number;
alter table ybook delete subNumber;

alter table yGuideDetail delete pId;

alter table yGuidePurpose delete number;
alter table yGuidePurpose delete subNumber;
alter table yGuidePurpose delete descript;

/*
alter table sproducts add constraint sproducts_856 foreign key (productid) references sguideproducts(prid) on update cascade;
alter table sguideproducts add constraint sguideproducts_857 foreign key (prSeriaId) references sguideseries(seriaid) on update cascade;
alter table sguideproducts add constraint sGuideFormulssGuideProducts foreign key (formulaNom) references sGuideFormuls(nomer) on update cascade;
alter table sguidenomenk add constraint sguideformulssguidenomenk foreign key (formulaNom) references sGuideFormuls(nomer) on update cascade;
alter table sguidenomenk add constraint sguideformulssguidenomenk1 foreign key (formulaNomW) references sGuideFormuls(nomer) on update cascade;
alter table sdocs add constraint sguidesourcesdocs foreign key (SourId) references sguidesource(sourceId) on update cascade;
alter table sdocs add constraint sguidesourcesdocs1 foreign key (DestId) references sguidesource(sourceId) on update cascade;
alter table sguideklass add constraint sguideklass_851 foreign key (parentklassId) references sguideklass(klassId) on update cascade;
alter table sguideseries add constraint sguideseries_858 foreign key (parentseriaId) references sguideseries(seriaId) on update cascade;
*/


alter table yBook modify UEsumm default 0;
alter table yBook modify Debit default     '255';
alter table yBook modify subDebit default  '00';
alter table yBook modify Kredit default    '255';
alter table yBook modify subKredit default '00';
alter table yBook modify KredDebitor default 0;
alter table yBook modify ordersNum default '';
alter table yBook modify purposeId default 0;
alter table yBook modify detailId default 0;
alter table yBook modify M default '';
alter table yBook modify NomerZap default '';
alter table yBook modify firm default '';
alter table yBook modify Note default '';


------------------------------------------------------------------------- 
--    Изменить тип поля счетов во всех таблицах подсистемы бухгалтерии с 
--	целого на символьный. Потому что в Комтехе эти поля символьные
------------------------------------------------------------------------- 

alter table yBook drop primary key;	
alter table yGuideDetail  drop primary key;
alter table yGuidePurpose drop primary key;
alter table yGuideSchets  drop primary key;

alter table ybook         modify Debit       char(26);
alter table ybook         modify Kredit      char(26);
alter table ybook         modify SubDebit    char(10);
alter table ybook         modify SubKredit   char(10);
alter table yGuideDetail  modify Debit       char(26);
alter table yGuideDetail  modify Kredit      char(26);
alter table yGuideDetail  modify SubDebit    char(10);
alter table yGuideDetail  modify SubKredit   char(10);
alter table yGuidePurpose modify Debit       char(26);
alter table yGuidePurpose modify Kredit      char(26);
alter table yGuidePurpose modify SubDebit    char(10);
alter table yGuidePurpose modify SubKredit   char(10);
alter table yGuideSchets  modify Number      char(26);
alter table yGuideSchets  modify SubNumber   char(10);

------------------------------------------------------------------------- 
--        Восстанавливаем первичные ключи 
------------------------------------------------------------------------- 

--alter table yBook          add primary key (xDate, Debit, SubDebit, Kredit, SubKredit);
alter table yGuideDetail   add primary key (Debit, SubDebit, Kredit, SubKredit, purposeId, id);
alter table yGuidePurpose  add primary key (Debit, SubDebit, Kredit, SubKredit, pId);
alter table yGuideSchets   add primary key (Number, SubNumber);

create unique index unique_purpose on yGuidePurpose (Debit, subDebit, Kredit, subKredit, pDescript);


-- денормализация "Уточнения" ЖХО
alter table ybook add descript varchar(50) default '';
alter table ybook add purpose varchar(50) default '';


-- Триггер, автоматически контролирующий (и корректирующей)
-- правильность заполнения yGuidePurpose.pId 

if exists (select 1 from systriggers where trigname = 'purposeId_bifer' and tname = 'yGuidePurpose') then 
	drop trigger yGuidePurpose.purposeId_bifer;
end if;

create TRIGGER purposeId_bifer before insert on
yGuidePurpose
referencing new as new_name
for each row
begin

	declare v_purposeid integer;
	declare v_debit_sc   varchar(26);
	declare v_debit_sub  varchar(10);
	declare v_credit_sc  varchar(26);
	declare v_credit_sub varchar(10);
	declare v_purpose    varchar(99);
	declare v_currentId integer;


	declare c_purposes dynamic scroll cursor for
		select pid from yGuidePurpose
		where 
				Debit = v_debit_sc
			and subDebit = v_debit_sub
			and Kredit = v_credit_sc
			and subKredit = v_credit_sub
		order by pId asc
	;
	


	set v_purposeid = new_name.pid;
	set v_purpose = new_name.pDescript;
	set v_debit_sc  = new_name.debit;
	set v_debit_sub = new_name.subDebit;
	set v_credit_sc  = new_name.kredit;
	set v_credit_sub = new_name.subKredit;

	if v_purposeid is null or v_purposeid = 0 then

	    -- требуется добавить такое Назначение
	    -- Сначала ищем первый свободный id включая дырки в последовательности
		set v_purposeId = 0;

		open c_purposes;

		find_id: loop
			fetch c_purposes into v_currentId;

			if SQLCODE != 0 then 
				--set v_purposeId = v_purposeId + 1;
				leave find_id;
			end if;

			if v_purposeId != v_currentId then
				leave find_id;
			end if;
			set v_purposeId = v_purposeId + 1;

		end loop;

		close c_purposes;

		
		if not exists (select 1 from yGuidePurp where descript = v_purpose) then
			insert into yGuidePurp (descript) values (v_purpose);
		end if;
			
		-- исправить поле на новое свобоное значение
		set new_name.pId = v_purposeId;
	end if;
end;




-- Денормализуем поле "Уточнение"
update ybook b set descript = d.descript from yguidedetail d 
where 
		d.id = b.detailid 
	and d.debit = b.debit 
	and d.subDebit = b.subDebit 
	and d.Kredit = b.Kredit 
	and d.subKredit = b.subKredit
	and b.purposeid = d.purposeid;

-- Переносим то, что было в  yGuideDetail в y GuidePurpose
-- Формируем "Назначение", по алгоритму. Если "Уточнение" не пустое, то теперь
-- оно должно попасть в "Назначение". Если же уточнения нет ( пустое), назначение
-- остается без изменения.
update ybook b 
	set purpose = (if d.descript is null or d.descript = '' then p.pDescript else d.descript endif
	)
from yguidedetail d, yguidepurpose p
where 
	d.id = b.detailid and d.debit = b.debit and d.subDebit = b.subDebit and d.Kredit = b.Kredit and d.subKredit = b.subKredit and b.purposeid = d.purposeid
	and p.debit = b.debit and p.subDebit = b.subDebit and p.Kredit = b.Kredit and p.subKredit = b.subKredit and b.purposeid = p.pid
;

-- все старые "уточнения" теперь в "назначении"
update ybook b set descript = null;

insert into yGuidePurpose (Debit, subDebit, Kredit, subKredit, pDescript)
select distinct Debit, subDebit, Kredit, subKredit, purpose
from ybook b
where not exists (select 1 from yGuidePurpose p 
where 
	p.debit = b.debit and p.subDebit = b.subDebit and p.Kredit = b.Kredit and p.subKredit = b.subKredit and b.purpose = p.pDescript
);


update ybook b set purposeId = p.pId 
from yguidepurpose p
where 
	p.debit = b.debit 
	and p.subDebit = b.subDebit 
	and p.Kredit = b.Kredit 
	and p.subKredit = b.subKredit
	and p.pDescript = b.purpose
	;



-- перенос настроек для "Реализации" из yGuideDetail в yGuidePurpose
alter table yGuidePurpose add auto varchar(1);
update yGuidePurpose set auto = '';
alter table yGuidePurpose modify [auto] null default '';

update yGuidePurpose b 
	set auto = d.auto
from yGuideDetail d 
where 
		d.debit = b.debit 
	and d.subDebit = b.subDebit 
	and d.Kredit = b.Kredit 
	and d.subKredit = b.subKredit
	and b.pid = d.purposeid
	and d.auto != '';

-- после денормализации удаляем ненужные столбы...
alter table ybook drop detailId;
alter table ybook drop purpose;

-- ... и теперь уже не используемые таблицы 
-- для нормализованных уточнений
drop table yguidedet;
drop table yguidedetail;



------------------------------------------------------------------------- 
--        Восстанавливаем внешние ключи 
------------------------------------------------------------------------- 

--alter table ybook add constraint yBook_381 foreign key (Debit, subDebit, Kredit, subKredit, purposeId, detailId) references yGuideDetail(Debit, subDebit, Kredit, subKredit, purposeId, id) on update cascade;
--alter table ybook add constraint yBook_382 foreign key (Debit, subDebit, Kredit, subKredit, purposeId) references yGuidePurpose(Debit, subDebit, Kredit, subKredit, pId) on update cascade;
alter table ybook add constraint yBook_384 foreign key (Debit, subDebit) references yGuideSchets(number, subNumber) on update cascade;
alter table ybook add constraint yBook_385 foreign key (Kredit, subKredit) references yGuideSchets(number, subNumber) on update cascade;
--alter table yGuideDetail add constraint yGuideDetail_383 foreign key (Debit, subDebit, Kredit, subKredit, purposeId) references yGuidePurpose(Debit, subDebit, Kredit, subKredit, pId) on update cascade;
--alter table yGuideDetail add constraint yGuideDetyGuideDetail foreign key (descript) references yGuideDet(descript) on update cascade;
alter table yGuidePurpose add constraint yGuidePurpose_386 foreign key (Debit, subDebit) references yGuideSchets(number, subNumber) on update cascade on delete cascade;
alter table yGuidePurpose add constraint yGuidePurpose_387 foreign key (Kredit, subKredit) references yGuideSchets(number, subNumber) on update cascade on delete cascade;
alter table yGuidePurpose add constraint yGuidePurpyGuidePurpose foreign key (pDescript) references yGuidePurp(descript) on update cascade;

-- Добавить лидирующие нули в номера счетов и субсчетов
delete from yguideschets where number = '' and subnumber = '';

update yGuideSchets set Number = cast_acc(Number);
update yGuideSchets set subNumber = cast_acc(subNumber);

-- Почему-то эти ограничения запрещают добавлять 
-- записи через Комтек. 
-- Хорошо бы разобраться в чем дело...
alter table ybook drop foreign key yBook_384;
alter table ybook drop foreign key yBook_385;


------------------------------------------------------------------------- 
-- ВАРИАНТНЫЕ ИЗДЕЛИЯ БЕЗ UNIKEY 
------------------------------------------------------------------------- 

	create table sGuideVariant (
		c int
		, productid int
		, xgroup char(1)
	);

	create table sVariantPower (
		numgroup int
		, productid smallint
		, fixgroups integer
		-- ссылается на корневую папку вариантного изделия.
		-- если ни одного варианта еще не "материализовались", то тогда 
		-- это поле равно null
		, id_inv integer 
	);

    create table sGuideComplect (
		id_variant int default autoincrement
		, Productid int not null
		, xPrExt int
		, id_inv integer
	);

	create table sVariantComplect (
		id_variant int
		, nomNom varchar(20)
		, id_compl integer
	);
 	


create table edizm (id_edizm integer, name varchar(100));
create table size( id_size integer, name varchar(100));

-- Фиксируем время трансдатации, которое 
-- впоследствии будет использоваться для 
-- контролирования условия обработки заказов
-- Менеджеры, которые будут заводить заказы 
-- после этой даты, будут обзязаны выставлять 
-- предприятие, через которую этот заказ будет
-- выполняться.


end if; 
	


if not exists (select 1 from systable where table_name = 'GuideVenture') then
create table GuideVenture (
	 ventureId integer default autoincrement
	,ventureName varchar(128)
	,sysname varchar(32)
	,invCode varchar(10)
	,standalone integer default 0
	,primary key (ventureId)
);

alter table orders add ventureId integer;

alter table orders add foreign key fk_venture (ventureId) references GuideVenture(ventureId);

insert into GuideVenture (ventureName, sysname, invCode)
values ('Петровские мастерские', 'pm', '50');

insert into GuideVenture (ventureName, sysname, invCode)
values ('Маркмастер', 'mm', '55');

insert into GuideVenture (ventureName, sysname, invCode)
values ('Аналитика', 'st', '');

-- фиктивный товар "услуга гравировки"
insert into sguidenomenk (nomnom, nomname, klassid) select 'УСЛ', 'Услуга гравировки', 0

end if;
