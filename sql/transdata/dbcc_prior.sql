/*
alter table guideVenture add intInvoice integer;

begin 
	declare v_max_nu_char varchar(100);
	declare v_max_nu integer;

	for all_serv as a dynamic scroll cursor for
		select srvname as v_slave_server from sys.sysservers 
	do
		set v_max_nu_char = select_remote(
			v_slave_server
			, 'jscet'
			, 'max(convert(integer, nu))'
			, 	  'isnumeric (nu) = 1'
			 	+ ' and convert(varchar(4), dat, 112) = convert(varchar(4), now(), 112)'
		);

		set v_max_nu = isnull(convert(integer, v_max_nu_char), 0) + 1;

		update guideVenture set intInvoice = v_max_nu where sysname = v_slave_server;

	end for;
end;


create table sBadge (
	productid smallint
	, nomnom varchar(20)
	, id_size integer
--	, classId integer
--	, nm varchar(32)
	, id_compl integer
	, primary key (productid, nomnom, id_size)
);


alter table sBadge add constraint b_product foreign key (productid, nomnom) references sproducts (productid, nomnom) on update cascade

alter table size modify id_size not null;

alter table size add primary key (id_size);

alter table sBadge add constraint fk_sz foreign key (id_size) references size (id_size) on update cascade;

alter table sProducts add isbadge tinyint;

update sproducts set isbadge = 0;

alter table sproducts modify isbadge not null default 0;


-- remove badges artefacts
drop table sBadge;

alter table sproducts delete isbadge;

-- ����� � �������
update sdmcrez set quantity = 2 where numdoc = 5092006 and nomnom = '1002L515'
commit

if exists (select '*' from sys.sysservers where srvname = 'priorOtl') then  
	drop server priorOtl;
end if;


--alter table bayorders add id_bill integer;

if exists(select 1 from SYS.SYSCOLUMNS WHERE CREATOR = 'DBA' AND TNAME = 'sproducts' and cname = 'isbadge') then
	alter table sproducts delete isbadge;
end if;
*/

/*
����������� �� ��������� ������������
select n.nomnom, w.nomnom, n.nomname, w.nomname, n.cod, w.cod, n.cost was, w.cost now from sguidenomenk n
join sguidenomenk_pwork w on n.id_inv = w.id_inv and n.cost != w.cost


select * from sguidenomenk where cod like 'CLT%'

select n.nomnom, w.nomnom, n.nomname, w.nomname, n.cod, w.cod, n.cost was, w.cost now from sguidenomenk n
join sguidenomenk_pwork w on n.nomnom = w.nomnom and n.cod != w.cod


-- ����� �������� ������ ����� ��� ���������� ���������� 
-- nomNom �� ����������� ������������
select * from sys.syscolumns
where cname like '%nomnom%'
*/



/*************************************************************\
*                                                             *
* ���������� ������� ������ ��� ������, ������������,         *
* ������������                                                *
* �������� ����� ������������ ����� �������������             *
* (��������, ��� ��������� ����������), ����� �����������     *
* ��������� ���������� �� ������� ������                      *
* ��� ����������� ������ ��� ���� �������, �� ���� ������     *
* ��� 3-� ������: xVariantnomenc, BayNomenkOut �              *
* ��� sVariantComplect                                        *
\************************************************************* /


-- ����� � ���������
delete xvariantnomenc
where not exists (
	select 1 from sguidenomenk n where xvariantnomenc.nomnom = n.nomnom
);


if exists(select * from SYS.SYSFOREIGNKEYS WHERE FOREIGN_CREATOR = 'DBA' AND ROLE = 'xvariantnomencnomnom') then
   alter table xVariantnomenc drop foreign key xvariantnomencnomnom;
end if;

alter table xVariantnomenc add foreign key xvariantnomencnomnom ( nomNom)  references sGuideNomenk ( nomNom)
on update cascade;



-- � ��������� �� �������� ��� ��������� �����
delete baynomenkout 
where not exists 
(select 1 from sdmcrez r where 
	r.numdoc = baynomenkout.numorder and r.nomnom = baynomenkout.nomnom
);



if exists(select * from SYS.SYSFOREIGNKEYS WHERE FOREIGN_CREATOR = 'DBA' AND ROLE = 'xPredmetyByNomenkBayNomenkOut') then
   alter table BayNomenkOut drop foreign key xPredmetyByNomenkBayNomenkOut;
end if;

alter table BayNomenkOut add foreign key xPredmetyByNomenkBayNomenkOut (numOrder,  nomNom)  references sDmcRez (numDoc, nomNom)
on update cascade on delete cascade;

if exists(select * from SYS.SYSFOREIGNKEYS WHERE FOREIGN_CREATOR = 'DBA' AND ROLE = 'xPredmetyByNomenkGuideNomenk') then
   alter table BayNomenkOut drop foreign key xPredmetyByNomenkGuideNomenk;
end if;
alter table BayNomenkOut add foreign key xPredmetyByNomenkGuideNomenk (nomNom)  references sGuideNomenk (nomNom)
on update cascade;


-- � ���� � ���������� ��������
delete sVariantComplect
where not exists (
	select 1 from sguidenomenk n where sVariantComplect.nomnom = n.nomnom
);

if exists(select * from SYS.SYSFOREIGNKEYS WHERE FOREIGN_CREATOR = 'DBA' AND ROLE = 'sVariantComplectnomnom') then
   alter table sVariantComplect drop foreign key sVariantComplectnomnom;
end if;

alter table sVariantComplect add foreign key sVariantComplectnomnom ( nomNom)  references sGuideNomenk ( nomNom)
on update cascade;
**********************************************************/


/*********************************************************\
*         ������ ������������ ����������� �������         *
* ��� ����������, ����������������� ��� ���������         *
* ����� ������������ ������� �������� ������������        *
\*********************************************************/

-- ���������� �������� ������������ � STime �� �����, ������ ���
-- ��� ���������� �� �� ���������� xPredmetyBy... � �� 
-- ���������� ������ sDmc/sDmcRez


-- ��-�� ����, ��� ���������� ������ �� ������� ������� ��� ������� ������ 
-- �������� �������� ������������ ������������� ���������
-- ����������� �������. ����� ����� ���������� :-(
-- ��� ����� - ������� ������� ���, ��� ���������������,
-- ����� ������ ���������.

-- ��������, ��� ��� ������ ���.
--select count(*), id_variant, nomnom from svariantcomplect group by id_variant, nomnom having count(*) > 1


/*
create table #tmp (id_variant integer, nomnom varchar(20));

insert into #tmp 
select id_variant, nomnom from svariantcomplect group by id_variant, nomnom having count(*) > 1;

-- todo 
--����������, ��� ��������� ������������ � ������������ (����) �����
delete from svariantcomplect c
from #tmp t
where c.id_variant = t.id_variant and c.nomnom = t.nomnom;


-- ���������� ������������ ���������� �������, ����������/������������
-- ����� ������������
begin
	
	declare v_id_compl integer;

	declare v_table_name varchar(100);
	declare v_fields varchar(1000);
	declare v_values varchar(1000);


	
	for spoiled_products as sp dynamic scroll cursor for
		select 
			id_variant as r_id_variant
			, gc.productid as r_productid
			, xprExt as r_xPrExt
			, id_inv as r_id_inv
		from sguidecomplect gc 
		join (select distinct gv.productid from sguidevariant gv where gv.xgroup = '' or gv.c = 1) gv 
			on gc.productid = gv.productid
	do
		-- ����������� � sVariantComplect
		for fixed_nomnom as fn dynamic scroll cursor for
			select p.nomNom as r_nomnom
			 	, n.id_inv as r_id_inv_compl
				, e.id_edizm as r_id_edizm
				, p.quantity as r_kol
			from sproducts p
			join sguidenomenk n on n.nomnom = r_nomnom and p.nomnom = n.nomnom
			join edizm e on e.name = n.ed_izmer
			where 
				p.productId = r_productid --576
			and exists (select 1 from sguidevariant vp where 
							p.xgroup = vp.xgroup 
							and vp.productid = p.productid and (vp.c = 1 or vp.xgroup = '')
						)
			and not exists (
					select 1 from svariantcomplect cg where cg.id_variant = r_id_variant and cg.nomnom = n.nomnom
				)
	
		do
			set v_id_compl = get_nextid('compl');
			
			set v_fields ='id'
			+ ', id_inv'
			+ ', id_inv_belong'
			+ ', id_edizm'
			+ ', kol'
			;
			set v_values =
				 convert(varchar(20), v_id_compl)
				+ ', ' + convert(varchar(20), r_id_inv_compl)
				+ ', ' + convert(varchar(20), r_id_inv)
				+ ', ' + convert(varchar(20), r_id_edizm)
				+ ', ''''' + convert(varchar(20), r_kol) + ''''''
			;	
	    
			call insert_host ('compl', v_fields, v_values);
	
					
			insert into svariantcomplect (id_variant, nomnom, id_compl)
			values (r_id_variant, r_nomnom, v_id_compl);
		end for; -- by nomnom
	end for;     -- by products
	
end;    -- of begin
*/


/*****************************************************
begin
	for nom as n dynamic scroll cursor for
		select 
			nomnom as r_nomnom, ed_izmer as p_izm1, ed_izmer2 as p_izm2
			, i.id_esizm1 as 
		from sguidenomenk n
		join edizm e on e.name = n.ed_izmer
		join inv_stime i on i.id = n.id_inv
		join edizm ie on ie.id_edizm = i.id_edizm1 and ie.id_edizm = e.id_edizm
		where perlist != 1
	do
		
	end for;
end;


select * from sguidenomenk n
join size s on s.name = n.size
join inv_stime i on i.id = n.id_inv
join size ie on ie.id_size = i.id_size and ie.name != s.name

*********************************************************/


/*

	-- ��������, ��� ��� ������������ � stime ����� ��������������� 
	-- ������������ � ���� prior. 
	-- ���� �� �� ��� �� ����� ������ �������� ��� ������������
	-- � ���� �������: ����������� exception
	-- ��� �������������� ����� ������������� ������ ������, 
	-- ������ ��� ��� ����� �� id_inv. �
	-- � ����� ������������ ����� ���� �������
	-- ����� ���� ������ �� ���� cod?

create existing table inv_stime at 'stime...inv';

select * from inv_stime i 
where 
not exists (select 1 from sguidenomenk n where i.id = n.id_inv)
and not exists (select 1 from sguideproducts n where i.id = n.id_inv)
and not exists (select 1 from sguideseries n where i.id = n.id_inv)
and not exists (select 1 from sguideklass n where i.id = n.id_inv)
and not exists (select 1 from sguidecomplect c where c.id_inv = i.id)
order by 1 desc;


drop table inv_stime;

-- ������������, ������� ��� ��������������� � Prior
-- ��������� ����������� ������������ ��������������� ������������
-- call delete_host('inv', 'id in (6820, 6825)');
*/


/*

**********************************************
������� � ������� ���� Prior 9 ������� 2006 
**********************************************

-- ���� �������� ������������ �� ������������

-- �������� �������� ��������� ���� � ���������� ���������
-- xDate ��� ���������� ���������

alter table sdocs delete primary key;
alter table sdocs add primary key (numDoc, numExt);

-- ������ ����� �������� ����������� � �� ������� xDmcxxx
alter table sDmc add constraint sDmcDoc foreign key (numDoc, numExt) references sDocs(numDoc, numExt) on delete cascade on update cascade;

alter table sDmcRez add constraint sDmcrezDoc foreign key (numDoc, numExt) references sDocs(numDoc, numExt) on delete cascade on update cascade;

alter table sDmcmov add constraint sDmcmovDoc foreign key (numDoc, numExt) references sDocs(numDoc, numExt) on delete cascade on update cascade;

-- � ������� ����� ��������� ������� ����, ��� ������ ��� �����������
-- �� �����������(�) �������� �� ����������.
-- �� ������� �������, ��� ��� ������ ��������� �� "���������� ����������"

if exists (select 1 from systable where table_name = 'sDocsIncome') then
	drop table sDocsIncome;
end if;

create table sDocsIncome (
	  numDoc integer
	, numExt tinyint
	, nomnom varchar(20)
	, ventureId integer
	, id_analytic integer
	, id_jmat integer
);

alter table sDocsIncome add constraint sDocsIncomeDoc foreign key (numDoc, numExt) references sDocs(numDoc, numExt) on update cascade on delete cascade;
alter table sDocsIncome add constraint sDocsIncomeVenture foreign key (ventureId) references guideVenture(ventureId) on update cascade on delete cascade;
alter table sDocsIncome add constraint sDocsIncomeNomnom foreign key (nomnom) references sGuideNomenk(nomnom) on update cascade on delete cascade;

-- � ���� Comtec.Stime ����� ����������� ������� � ��������� ���������
-- 
if not exists (select 1 from sys.syscolumns where tname = 'guideventure' and cname = 'id_analytic') then
	alter table guideventure add id_analytic integer;
end if;

if not exists (select 1 from sys.syscolumns where tname = 'system' and cname = 'id_analytic_default') then
	alter table system add id_analytic_default integer;
end if;

begin 
	declare v_id integer;
	declare default_income_text varchar(100);
	declare c_ventureName varchar(30);

	set c_ventureName = '��';
	set default_income_text = '''''������ �� '+ c_ventureName +'''''';
	set v_id = select_remote('stime', 'analytic', 'id', 'code = ' + default_income_text);
	if v_id is null then
		set v_id = insert_count_remote('stime', 'analytic', 'code', default_income_text);
	end if;
	update system set id_analytic_default = v_id;
	update guideventure set id_analytic = v_id where ventureName = '���������� ����������' ;

	set c_ventureName = '��';
	set default_income_text = '''''������ �� '+ c_ventureName +'''''';
	set v_id = select_remote('stime', 'analytic', 'id', 'code = ' + default_income_text);
	if v_id is null then
		set v_id = insert_count_remote('stime', 'analytic', 'code', default_income_text);
	end if;
	update guideventure set id_analytic = v_id where ventureName = '����������' ;
end;

if exists (select 1 from systable where table_name = 'sDocsVenture') then
	drop table sDocsVenture;
end if;

create table sDocsVenture (
	  id integer default autoincrement
	-- ���� �������� ��������� ������������
	, xDate datetime not null default current timestamp
	-- ���� ������������� ��������� ������������
	, nDate date
	-- ������� ������ ������������: ������
	, termFrom date
	-- ������� ������ ������������: �����
	, termTo   date null default current date
	-- �����, ����� ���� ������� ����������� 
	-- ��� ���� ������������� ����������� ��������� 
	-- ��� ��������� ������� � ����� ������ ��������
	, calculatedDatetime datetime
	-- �����, ������� �������� ���������
	, srcVentureId integer
	-- �����, ���������� ���������
	, dstVentureId integer
	-- ���������
	, note varchar(150)
	-- ������� ���������� ������������� ����������
	, procent real
	-- ������ �� ��������� � ���� �������
	, id_jmat integer
	, primary key(id)
);

alter table sDocsVenture add constraint srcFK foreign key (srcVentureId) references GuideVenture(ventureId);
alter table sDocsVenture add constraint dstFK foreign key (dstVentureId) references GuideVenture(ventureId);

if exists (select 1 from systable where table_name = 'sDmcVenture') then
	drop table sDmcVenture;
end if;

-- ������� ����������� ������������ �� ������������
create table sDmcVenture (
      -- ��������� ����
	  id integer default autoincrement
	  -- � ������ ��
	, sdv_id integer
	-- ����� ������������
	, nomnom varchar(20)
	-- ���������� ���������
	, quant float
	-- ���� ��� ��������
	, costed float
	-- ������ �� ������ � ���� �������
	, id_mat integer
);

alter table sDmcVenture add constraint docFK foreign key (sdv_id) references sDocsVenture(id) on update cascade on delete cascade;
alter table sDmcVenture add constraint nomnomFK foreign key (nomnom) references sGuideNomenk(nomnom) on update cascade;

--insert into sDocsVenture (srcVentureId, dstVentureId, termTo) select 1, 2, convert(date, '20051013') union select 2, 1, convert(date, '20051013');


commit;


// ����������� �� 29 ������ 2006 ��� ������������� ����� ������. �� ����������
if not exists (select 1 from sys.syscolumns where tname = 'guideventure' and cname = 'rusAbbrev') then
	alter table guideventure add RusAbbrev varchar(10);

	update guideventure set RusAbbrev = '��' where sysname = 'accountN';
	update guideventure set RusAbbrev = '��' where sysname = 'markmaster';
	update guideventure set RusAbbrev = '��' where sysname = 'stime';
	commit;
end if;


begin atomic
	declare v_ventureid integer;
	declare v_id_analytic integer;

	declare c_docs dynamic scroll cursor for
		select ventureId, id_analytic 
	from guideventure
	where sysname = 'markmaster';

	open c_docs;
	aaa: loop
		fetch c_docs into v_ventureid, v_id_analytic;
		leave aaa;
	end loop;
	close c_docs;
	
	insert into sDocsIncome (numdoc, numExt, ventureId, id_analytic)
		  select 4322103, 255, v_ventureid, v_id_analytic 
	union select 5220705, 255, v_ventureid, v_id_analytic 
	union select 5272112, 255, v_ventureid, v_id_analytic;
	
--	update sdocsincome set ventureid = v_ventureid, id_analytic = v_id_analytic where ventureid is null;
	
end;


if not exists (select 1 from sys.syscolumns where tname = 'sdocsventure' and cname = 'cumulative_id') then 
	alter table sdocsventure add cumulative_id integer null default 0;

	alter table sdocsventure add constraint sdv_cumulative foreign key(cumulative_id) references sDocsVenture(id) on update cascade on delete cascade;

	create index sdv_cumulative_id on sDocsVenture(cumulative_id);

	alter table sdocsventure modify srcVentureId not null;

	alter table sdocsventure modify dstVentureId not null;

	alter table sdocsventure modify termFrom null;

	alter table sdocsventure modify termTo null default null;

	alter table sdocsventure modify nDate not null;
end if;


if not exists (select 1 from sys.syscolumns where tname = 'system' and cname = 'ivo_procent') then 
	alter table system add ivo_procent float;
	update system set ivo_procent = 10.0;
end if;
*/

--**********************************************
-- �������� ����������� � ������������ ������
--**********************************************
/*
-- ������� � ������� ���� Prior 08 ������ 2006 ����
if not exists (select 1 from systable where table_name = 'GuideCurrency') then 
	create table GuideCurrency (
		  currency_iso varchar(20)
		, id_currency integer
		, name varchar(50)
		, id_guide integer
		, tp1 integer
		, tp2 integer
		, tp3 integer
		, tp4 integer
		, primary key (currency_iso)
	)
	;
end if;

if not exists (select 1 from sys.syscolumns where tname = 'sGuideSource' and cname = 'currency_iso') then 
	alter table sGuideSource add currency_iso varchar(20);
	alter table sGuideSource add constraint sourceCurrency foreign key (currency_iso) references GuideCurrency(currency_iso);
end if;


begin
	declare v_fields varchar(1000);
	declare v_values varchar(2000);
	declare v_where varchar(1000);
	declare v_id_cur integer;

	set v_id_cur = select_remote('stime', 'currency', 'id', 'nm = ''''�����''''' );

	insert into GuideCurrency (currency_iso, id_currency, id_guide, tp1, tp2, tp3, tp4) 
	values ('RUR', v_id_cur, 1120, 1, 1, 2, 0);

	set v_id_cur = get_nextid ('currency');
	message v_id_cur to client;
	

	set v_fields = 'id, nm, base_1, base_2, base_0, sub_1, sub_2, sub_0, rem, iso_code';
	set v_values = 
		convert(varchar(20), v_id_cur)
		 +', ''''������ ���'''''
		 +', ''''������'''''
		 +', ''''�������'''''
		 +', ''''��������'''''
		 +', ''''����'''''
		 +', ''''�����'''''
		 +', ''''������'''''
		 +', ''''������ ���'''''
		 +', ''''USD'''''
	;

	call insert_remote('stime','currency'
--	call insert_host('currency'
		, v_fields
		, v_values
	);

	insert into GuideCurrency (currency_iso, id_currency, id_guide, tp1, tp2, tp3, tp4) 
	values ('USD', v_id_cur, 1127, 1, 1, 2, 7);


	set v_id_cur = v_id_cur + 1;
	

	set v_fields = 'id, nm, base_1, base_2, base_0, sub_1, sub_2, sub_0, rem, iso_code';
	set v_values = 
		convert(varchar(20), v_id_cur)
		 +', ''''����'''''
		 +', ''''����'''''
		 +', ''''����'''''
		 +', ''''����'''''
		 +', ''''��������'''''
		 +', ''''���������'''''
		 +', ''''����������'''''
		 +', ''''����'''''
		 +', ''''EUR'''''
	;

	call insert_remote('stime','currency'
--	call insert_host('currency'
		, v_fields
		, v_values
	);

	insert into GuideCurrency (currency_iso, id_currency, id_guide, tp1, tp2, tp3, tp4) 
	values ('EUR', v_id_cur, 1127, 1, 1, 2, 7);

end;

*/


/*
--
-- ������� � �������� 10 ������ 2006 ����
-- ��� ��� ����������� 22 ������ 2006 ����

begin
	declare v_id_guide integer;
	declare v_tp1 integer;
	declare v_tp2 integer;
	declare v_tp3 integer;
	declare v_tp4 integer;

	for all_inc as a dynamic scroll cursor for
		select d.id_jmat as r_id_jmat
			, isnull(c.id_guide, ru.id_guide) as r_id_guide
			, isnull(c.id_currency, ru.id_currency) as r_id_currency
		from sguideSource s
		join GuideCurrency ru on ru.currency_iso = 'RUR'
		left join GuideCurrency c on c.currency_iso = s.currency_iso
		join sdocs d on d.id_jmat is not null and d.numExt = 255 and d.sourId = s.sourceId and xDate >= convert(varchar(20), '20060401')
	do
		message r_id_jmat to client;
		call gualify_guide(r_id_guide, v_tp1, v_tp2, v_tp3, v_tp4);
		call order_import_stime(
			r_id_jmat
			, r_id_currency
			, r_id_guide
			, v_tp1
			, v_tp2
			, v_tp3
			, v_tp4
		);
	end for;
end;
*/


/*
// ��������� �-�� �� ��������� ���������
begin
	for all_inc as a dynamic scroll cursor for
		select 
			m.numdoc as r_numdoc, m.numext as r_numext 
			, m.id_mat as r_id_mat, m.quant as r_quant
			, n.perList as r_perList
		from sdocs d
		join sdmc m on m.numdoc = d.numdoc and m.numext = d.numext
		join sguidenomenk n on n.nomnom = m.nomnom
		where 
			m.numext = 255
			and m.id_mat is not null
        order by m.id_mat
	do
        message r_id_mat, ',',r_quant, ',',r_perlist to client;
		call change_mat_qty_stime(r_id_mat, r_quant/r_perList);
	end for;
end;
*/


/********** ����������������� ����� ************

if exists (select 1 from systable where table_name = 'sPriceHistory') then
	drop table sPriceHistory;
end if;

create table sPriceHistory (
	  nomnom varchar(20)    // ����� ������������
	, cost float            // ����, � ������� ����. ������� �������� � sGuideNomenk.cena1
	, change_date datetime  // ����� ���������
	, changed_by_id tinyint // ��� ��������
);

alter table sPriceHistory add constraint price_history foreign key (nomnom) references sGuideNomenk(nomnom) on update cascade on delete cascade;
alter table sPriceHistory add constraint changed_by foreign key (changed_by_id) references GuideManag(managId) on update cascade on delete cascade;

-- � ������� � 24 ��� 2006
if not exists (select 1 from systable where table_name = 'sPriceBulkChange') then
	create table sPriceBulkChange (
		id integer default autoincrement
		, xDate datetime default current timestamp
		, guide_klass_id smallint null
		, changed_by tinyint
		, primary key (id)
	);

	alter table sPriceHistory add bulk_id integer;

	alter table sPriceHistory add constraint bulk_change foreign key (bulk_id) references sPriceBulkChange(id) on delete cascade;

	alter table sPriceBulkChange add constraint guide_klass foreign key (guide_klass_id) references sGuideKlass(klassId) on delete cascade;

end if;
*/



/*
--- ����������� ��� ������� ������������� 
-- ������� � �������� 1 ���� 2006 ����

-- ��������� �������������� ������
begin
	for all_inc as a dynamic scroll cursor for
		select nomnom as r_nomnom, id_inv as r_id_Inv from sguidenomenk
	do
		call update_host('inv', 'nomen', '''''' + r_nomnom + '''''', 'id = ' + convert(varchar(20), r_id_inv));
	end for;
end;


-- � ��������� ��������� "�� ����" � "����" � ������������ (?) ���������
begin
	for all_inc as a dynamic scroll cursor for
		select 
			d.id_jmat
			, src.id_voc_names id_src
			, dst.id_voc_names id_dst
		from sdocs d
		join sguidesource src on src.sourceId = d.sourid
		join sguidesource dst on dst.sourceId = d.destid 
		where numext = 254
		and id_jmat is not null
		order by xdate 
	do
		call update_remote('stime', 'jmat','id_s', convert(varchar(20), id_src), 'id = ' + convert(varchar(20), id_jmat));
		call update_remote('stime', 'jmat','id_d', convert(varchar(20), id_dst), 'id = ' + convert(varchar(20), id_jmat));
	end for;
end;


-- � ��������� ��������� ������ ������������ ��������� "�� ����" � "����" 
begin
	for all_inc as a dynamic scroll cursor for
		select 
			d.id_jmat
			, src.id_voc_names id_src
			, dst.id_voc_names id_dst
		from sdocs d
		join sguidesource src on src.sourceId = d.sourid
		join sguidesource dst on dst.sourceId = d.destid 
		where numext = 254
		and id_jmat is not null
		order by xdate 
	do
		call update_remote('stime', 'jmat','id_s', convert(varchar(20), id_src), 'id = ' + convert(varchar(20), id_jmat));
		call update_remote('stime', 'jmat','id_d', convert(varchar(20), id_dst), 'id = ' + convert(varchar(20), id_jmat));
	end for;
end;


-- � ���� ��������� ������ ������������ ��������� "�� ����" � "����" 
begin
	for all_inc as a dynamic scroll cursor for
		select 
			m.numdoc as r_numdoc, m.numext as r_numext 
			, m.id_mat as r_id_mat, m.quant as r_quant
			, n.perList as r_perList
		from sdocs d
		join sdmc m on m.numdoc = d.numdoc and m.numext = d.numext
		join sguidenomenk n on n.nomnom = m.nomnom
		where 
			--m.numext = 254 and 
			m.id_mat is not null
        order by m.id_mat
	do
        message r_id_mat, ',',r_quant, ',',r_perlist to client;
		call change_mat_qty_stime(r_id_mat, r_quant/r_perList);
	end for;
end;
*/
  


/*

-- ������������ ��� ��������� ������������, � ������� ��� ��������� ������� ����
-- ������� � �������� 1 ���� 2006 ����
alter table sPriceHistory drop foreign key  price_history;
alter table sPriceHistory add constraint price_history foreign key (nomnom) references sGuideNomenk(nomnom) on update cascade on delete cascade;
alter table sPriceHistory drop foreign key  changed_by;
alter table sPriceHistory add constraint changed_by foreign key (changed_by_id) references GuideManag(managId) on update cascade on delete set null;

if not exists(select 1 from sys.syscolumns where creator = 'dba' and tname = 'guideventure' and cname = 'activity_start') then
	alter table GuideVenture add activity_start date null;
	update guideventure set activity_start = '20041116' where sysname = 'markmaster';
end if;


if exists(select 1 from sys.syscolumns where creator = 'dba' and tname = 'sdocsventure' and cname = 'calculatedDatetime') then
	alter table sdocsventure drop calculatedDatetime;
end if;

if not exists(select 1 from sys.syscolumns where creator = 'dba' and tname = 'sdocsventure' and cname = 'invalide') then
	alter table sdocsventure add invalid integer null;
end if;
*/


begin 
	declare v_id integer;
	declare default_income_text varchar(100);
	declare c_ventureName varchar(30);

	set c_ventureName = '��';
	set default_income_text = '''''������ => ���������''''';
	set v_id = select_remote('stime', 'analytic', 'id', 'code = ' + default_income_text);
	if v_id is null then
		set v_id = insert_count_remote('stime', 'analytic', 'code', default_income_text);
	end if;
	update guideventure set id_analytic = v_id, activity_start = '20051013'  where sysname = 'stime' ;
end;

if not exists(select 1 from sys.syscolumns where creator = 'dba' and tname = 'orders' and cname = 'zalog') then
	alter table orders add zalog float null;
end if;

if not exists(select 1 from sys.syscolumns where creator = 'dba' and tname = 'orders' and cname = 'nal') then
	alter table orders add nal float null;
end if;

commit;
