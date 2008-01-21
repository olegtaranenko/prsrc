

if exists (select 1 from sysprocedure where proc_name = 'wf_breadcrump') then
	drop procedure wf_breadcrump;
end if;


if exists (select 1 from sysprocedure where proc_name = 'wf_breadcrump_klass') then
	drop procedure wf_breadcrump_klass;
end if;

create function wf_breadcrump_klass(
	  p_id integer
	, p_table_name varchar(64) default null
	, p_pk_column_name varchar(64) default null
	, p_text_column_name varchar(64) default null
) returns varchar(256)
begin
	declare v_parentid integer;
	declare v_parent_name varchar(64);

	if isnull(p_id, 0) = 0 then	
		set wf_breadcrump_klass = null;
		return;
	end if;

	select klassname, parentklassid into wf_breadcrump_klass, v_parentid from sguideklass where klassid = p_id;
	if isnull(v_parentid, 0) != 0 then
			set wf_breadcrump_klass =  wf_breadcrump_klass(v_parentid, p_table_name, p_pk_column_name, p_text_column_name) + ' / ' + wf_breadcrump_klass;
	end if;

end;


----------------------------------------------------------------------
-- Детализация по проданной номенклатуере с сортировкой по дереву
----------------------------------------------------------------------


if exists (select 1 from sysprocedure where proc_name = 'wf_sort_klasses') then
	drop procedure wf_sort_klasses;
end if;


CREATE procedure wf_sort_klasses(
)
begin
	declare v_lvl integer;
	declare v_prev_count integer;
	declare v_cur_pos integer;
	declare v_exit    integer;
	declare v_prev_id integer;
--create variable @v_lvl        integer;
--create variable @v_prev_count integer;
--create variable @v_cur_pos    integer;
--create variable @v_exit       integer;
--create variable @v_prev_id    integer;



	set v_lvl = 1;
	set v_prev_count = 0;
	set v_cur_pos = 1;

create table #klass_childs (parent integer, child_count integer, lvl integer);
--create table #klass_ordered (id integer, ord integer);

	insert into #klass_childs 
	select parentklassid, count(*), v_lvl from sguideklass 
	where parentklassid != 0
	group by parentklassid;

		
	branch: loop
		insert into #klass_childs
			select c.parentklassid, sum(child_count), v_lvl + 1
			from sguideklass c
			join #klass_childs p on c.klassid = p.parent
            where lvl = v_lvl and parentklassid != 0
			group by c.parentklassid;

		if @@rowcount = 0 then
			set v_lvl = v_lvl + 1;
			leave branch;
		end if;
		set v_lvl = v_lvl + 1;
	end loop;


	insert into #klass_childs
		select parent, sum(child_count), v_lvl + 1
		from #klass_childs p 
		group by parent;
	-- на уровне v_lvl + 1 хранятся полное количество детей включая всех потомков.



	-- нумеруем первый ряд после корня.
	for schich0 as s0 dynamic scroll cursor for
		select klassid as r_klassid, klassname as r_klassname, isnull(p.child_count, 0) as r_child_count
        from sguideklass k
        left join #klass_childs p on p.parent = k.klassid and p.lvl = v_lvl + 1
        where parentklassid = 0 and klassid != 0 order by klassname
	do
    
        insert into #klass_ordered (id, ord) select r_klassid, v_cur_pos;
        --message '[', r_klassid, '] ', r_klassname, ': ', v_cur_pos to client;
        set v_prev_count = r_child_count;
        set v_cur_pos = v_cur_pos + 1 + v_prev_count;

    end for;


	-- нумеруем оставшиеся ряды (или уровни).
	branch2: loop
        set v_exit = 0;

        set v_prev_count = 0;
        set v_prev_id = 0;
		for schich1 as s1 dynamic scroll cursor for
            select klassid as r_klassid, klassname as r_klassname, parentklassid as r_parentid, ord as r_ord, isnull(p.child_count, 0) as r_childs
            from sguideklass k
            join #klass_ordered o on o.id = k.parentklassid
            left join #klass_childs p on k.klassid = p.parent and p.lvl = v_lvl + 1
            where not exists (select 1 from #klass_ordered o1 where o1.id = k.klassid)
            order by parentklassid, klassname
		do
			if r_parentid != v_prev_id then
		        set v_cur_pos = r_ord + 1;
--                message v_cur_pos , ': ', r_childs, ': ', r_ord to client;
		    else 
		    	set v_cur_pos = v_cur_pos + 1 + v_prev_count;
		    end if;
    
            insert into #klass_ordered (id, ord) select r_klassid, v_cur_pos;
            set v_exit = 1;
--            message '[', r_klassid, '] ', r_klassname, '-', r_parentid, ': ', v_cur_pos, ': ', r_ord to client;
            set v_prev_id = r_parentid;
            set v_prev_count = r_childs;

    
        end for;
        if v_exit = 0 then 
        	leave branch2;
        end if;
	end loop;

	
--	select * from #klass_childs where lvl = v_lvl + 1;
--	select * from #klass_ordered order by ord; 

--	drop table #klass_ordered;
	drop table #klass_childs;

end;



if exists (select 1 from sysprocedure where proc_name = 'wf_nomenk_saled') then
	drop procedure wf_nomenk_saled;
end if;

CREATE procedure wf_nomenk_saled(
	  p_start datetime default null
	, p_end datetime default null
)
begin
	create table #nomenk_saled (nomnom varchar(20), quant double, cenaTotal double);
	create table #klass_ordered (id integer, ord integer);

	call wf_sort_klasses;


insert into #nomenk_saled
select d.prnomnom, sum(round(d.quant, 2)) as quant, sum(round(d.cenaEd, 2) * round(d.quant, 2))
from vPredmetyOutDetail d
WHERE d.type = 8 and outDate between isnull(p_start, '20010101') and isnull(p_end, '21001231')
group by d.prnomnom
;

select '' as outtype, trim(n.cod + ' ' + nomname + ' ' + n.size) as name, s.quant, s.cenaTotal, s.nomnom, o.ord
, k.klassid, wf_breadcrump_klass(k.klassid) as klassname, n.cost, n.ed_izmer2
from #nomenk_saled s
join sguidenomenk n on n.nomnom = s.nomnom
join #klass_ordered o on o.id = n.klassid
join sguideklass k on k.klassid = n.klassid
order by o.ord, 1
;

	drop table #nomenk_saled;
	drop table #klass_ordered;

end;


if exists (select 1 from sysprocedure where proc_name = 'wf_sale_nomenk_qty') then
	drop function  wf_sale_nomenk_qty;
end if;

create 
	function wf_sale_nomenk_qty(
	  p_nomnom varchar(20)
	, p_start datetime default null
	, p_end datetime default null
) returns double
begin

select sum(i.quant / n.perList) AS quant
    into wf_sale_nomenk_qty
    from BayOrders o 
    join sDocs d on d.numDoc = o.numOrder 
    join sDMC i on d.numExt = i.numExt and d.numDoc = i.numDoc
    join sDMCrez r on i.nomNom = r.nomNom and o.numOrder = r.numDoc
    join sGuideNomenk n ON n.nomNom = i.nomNom and r.nomNom = i.nomNom 
WHERE 
	    i.nomnom = p_nomnom
    and xDate between isnull(p_start, '20010101') and isnull(p_end, '21001231')
    ;

end;
--=============================================================================




if exists (select 1 from sysprocedure where proc_name = 'wf_sort_series') then
	drop procedure wf_sort_series;
end if;


CREATE procedure wf_sort_series(
)
begin
	declare v_lvl integer;
	declare v_prev_count integer;
	declare v_cur_pos integer;
	declare v_exit    integer;
	declare v_prev_id integer;
--create variable @v_lvl        integer;
--create variable @v_prev_count integer;
--create variable @v_cur_pos    integer;
--create variable @v_exit       integer;
--create variable @v_prev_id    integer;



	set v_lvl = 1;
	set v_prev_count = 0;
	set v_cur_pos = 1;

create table #seria_childs (parent integer, child_count integer, lvl integer);
--create table #seria_ordered (id integer, ord integer);

	insert into #seria_childs 
	select parentseriaid, count(*), v_lvl from sguideseries 
	where parentseriaid != 0
	group by parentseriaid;

		
	branch: loop
		insert into #seria_childs
			select c.parentseriaid, sum(child_count), v_lvl + 1
			from sguideseries c
			join #seria_childs p on c.seriaid = p.parent
            where lvl = v_lvl and parentseriaid != 0
			group by c.parentseriaid;

		if @@rowcount = 0 then
			set v_lvl = v_lvl + 1;
			leave branch;
		end if;
		set v_lvl = v_lvl + 1;
	end loop;


	insert into #seria_childs
		select parent, sum(child_count), v_lvl + 1
		from #seria_childs p 
		group by parent;
	-- на уровне v_lvl + 1 хранятся полное количество детей включая всех потомков.



	-- нумеруем первый ряд после корня.
	for schich0 as s0 dynamic scroll cursor for
		select seriaid as r_seriaid, serianame as r_serianame, isnull(p.child_count, 0) as r_child_count
        from sguideseries k
        left join #seria_childs p on p.parent = k.seriaid and p.lvl = v_lvl + 1
        where parentseriaid = 0 and seriaid != 0 order by serianame
	do
    
        insert into #seria_ordered (id, ord) select r_seriaid, v_cur_pos;
        --message '[', r_seriaid, '] ', r_serianame, ': ', v_cur_pos to client;
        set v_prev_count = r_child_count;
        set v_cur_pos = v_cur_pos + 1 + v_prev_count;

    end for;


	-- нумеруем оставшиеся ряды (или уровни).
	branch2: loop
        set v_exit = 0;

        set v_prev_count = 0;
        set v_prev_id = 0;
		for schich1 as s1 dynamic scroll cursor for
            select seriaid as r_seriaid, serianame as r_serianame, parentseriaid as r_parentid, ord as r_ord, isnull(p.child_count, 0) as r_childs
            from sguideseries k
            join #seria_ordered o on o.id = k.parentseriaid
            left join #seria_childs p on k.seriaid = p.parent and p.lvl = v_lvl + 1
            where not exists (select 1 from #seria_ordered o1 where o1.id = k.seriaid)
            order by parentseriaid, serianame
		do
			if r_parentid != v_prev_id then
		        set v_cur_pos = r_ord + 1;
--                message v_cur_pos , ': ', r_childs, ': ', r_ord to client;
		    else 
		    	set v_cur_pos = v_cur_pos + 1 + v_prev_count;
		    end if;
    
            insert into #seria_ordered (id, ord) select r_seriaid, v_cur_pos;
            set v_exit = 1;
            message '[', r_seriaid, '] ', r_serianame, '-', r_parentid, ': ', v_cur_pos, ': ', r_ord to client;
            set v_prev_id = r_parentid;
            set v_prev_count = r_childs;

    
        end for;
        if v_exit = 0 then 
        	leave branch2;
        end if;
	end loop;

	
--	select * from #seria_childs where lvl = v_lvl + 1;
--	select * from #seria_ordered order by ord; 

--	drop table #seria_ordered;
	drop table #seria_childs;

end;





if exists (select 1 from sysprocedure where proc_name = 'wf_nomenk_reliz') then
	drop procedure wf_nomenk_reliz;
end if;

CREATE procedure wf_nomenk_reliz(
	  p_start datetime default null
	, p_end datetime default null
)
begin
	create table #nomenk_reliz (nomnom varchar(20), quant double, sm double);
	create table #izdelia_reliz (prid integer, quant double, sm double, costTotal double);
	create table #klass_ordered (id integer, ord integer);
	create table #seria_ordered (id integer, ord integer);

	call wf_sort_series;
	call wf_sort_klasses;


insert into #izdelia_reliz
select 
	po.prId, sum(po.quant) as quant, sum(p.cenaEd * po.quant) as cenaTotal, sum(io.costEd * po.quant) as costTotal
from xpredmetybyizdeliaout po
join xpredmetybyizdelia p on p.numorder = po.numorder and p.prid = po.prid and p.prext = po.prext
join vIzdeliaOutPrimeCost io on po.outdate = io.outdate and io.numorder = po.numorder and io.prid = po.prid and io.prext = po.prext
WHERE po.outDate between isnull(p_start, '20010101') and isnull(p_end, '21001231')
group by po.prid
;


insert into #nomenk_reliz
select po.nomnom, sum(po.quant / n.perlist) as quant, sum(p.cenaEd * po.quant) as sum
from xpredmetybynomenkout po
join xpredmetybynomenk p on p.numorder = po.numorder and p.nomnom = po.nomnom
join sguidenomenk n on n.nomnom = po.nomnom and n.nomnom = p.nomnom
WHERE po.outDate between isnull(p_start, '20010101') and isnull(p_end, '21001231')
group by po.nomnom
;



select 'Изделия' as outtype, o.ord, trim(g.prDescript + ' ' + g.prsize) as name
	, quant, sm as cenaTotal, convert(varchar(20), r.prid) as id, g.prname as nomnom, g.prSeriaId as klassid
	, wf_breadcrump_seria(g.prseriaid) as klassname, 'шт.' as ed_izmer2
	, r.costTotal / quant as cost -- costEd
from #izdelia_reliz r
join sguideproducts g on g.prid = r.prid
join #seria_ordered o on o.id = g.prSeriaId
	union
select 'Номенклатура' as outtype, o.ord, trim(n.cod + ' ' + nomname + ' ' + n.size) as name
	, s.quant, s.sm as cenaTotal, s.nomnom as id, s.nomnom, k.klassid
	, wf_breadcrump_klass(k.klassid) as klassname, n.ed_izmer2
	, n.cost
from #nomenk_reliz s
join sguidenomenk n on n.nomnom = s.nomnom
join #klass_ordered o on o.id = n.klassid
join sguideklass k on k.klassid = n.klassid
order by 1, 2, 3
;


	drop table #nomenk_reliz;
	drop table #izdelia_reliz;
	drop table #klass_ordered;
	drop table #seria_ordered;

end;




if exists (select 1 from sysprocedure where proc_name = 'wf_breadcrump_seria') then
	drop procedure wf_breadcrump_seria;
end if;

create function wf_breadcrump_seria(
	  p_id integer
	, p_table_name varchar(64) default null
	, p_pk_column_name varchar(64) default null
	, p_text_column_name varchar(64) default null
) returns varchar(256)
begin
	declare v_parentid integer;
	declare v_parent_name varchar(64);

	if isnull(p_id, 0) = 0 then	
		set wf_breadcrump_seria = null;
		return;
	end if;

	select serianame, parentseriaid into wf_breadcrump_seria, v_parentid from sguideseries where seriaid = p_id;
	if isnull(v_parentid, 0) != 0 then
			set wf_breadcrump_seria =  wf_breadcrump_seria(v_parentid, p_table_name, p_pk_column_name, p_text_column_name) + ' / ' + wf_breadcrump_seria;
	end if;

end;



----------------------------------------------------------------------
--------------             sdmc      ------------------------
----------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_sdmc_income_bu' and tname = 'sdmc') then 
	drop trigger sdmc.wf_sdmc_income_bu;
end if;

create 
	trigger wf_sdmc_income_bu before update on 
sdmc
referencing new as new_name old as old_name
for each row
begin
	declare v_id_mat integer;
	declare v_perList float;
	declare v_curr float;
	declare v_id_jmat integer;
	declare v_summa float;
	declare v_summav float;
	declare v_quant float;
	declare v_sysname varchar(32);

	if update(quant) --and isnull(new_name.quant, 0.0) != isnull(old_name.quant, 0.0) 
	then
		set v_id_mat = old_name.id_mat;
		select perList into v_perList from sGuideNomenk where nomnom = old_name.nomNom;
		if v_id_mat is not null then
			select id_jmat, v.sysname into v_id_jmat, v_sysname
			from sdocs d
			left join guideventure v on v.ventureid = d.ventureid
			where d.numdoc = old_name.numdoc and d.numext = old_name.numext;
			set v_quant = new_name.quant/v_perList;

			call change_mat_qty_dual(v_sysname, v_id_mat, v_quant);
		end if;
	end if;
/*
	-- Не нужно, потому что в интерфейсе stime нельзя заменить номенклатуру.
	-- Можно только сначала удалить позицию, а потом завести новую.
	if update (nomnom) and isnull(new_name.nomnom, '') != isnull(old_name.nomnom, '') then
		set v_id_mat = old_name.id_mat;
		if v_id_mat is not null then
			select id_inv into v_id_inv from sguidenomenk where nomnom = new_name.nomnom;
			if v_id_inv is not null then
				call update_remote('stime', 'mat', 'id_inv', '''''' + v_id_inv + '''''', 'id = ' + convert(varchar(20), v_id_mat)); 
			end if;
		end if;
	end if;
*/
end;


if exists (select 1 from systriggers where trigname = 'wf_delete_sdmc' and tname = 'sdmc') then 
	drop trigger sdmc.wf_delete_sdmc;
end if;

if exists (select 1 from systriggers where trigname = 'wf_sdmc_bd' and tname = 'sdmc') then 
	drop trigger sdmc.wf_sdmc_bd;
end if;

create TRIGGER wf_sdmc_bd before delete on
sdmc
referencing old as old_name
for each row
begin
	declare remoteServer varchar(32);
	declare no_echo integer;
	set no_echo = 0;


  	begin
--  		message '@stime_sdmc = ', @stime_sdmc to log;
		select @stime_sdmc into no_echo; 
	exception 
		when other then
--			message 'Exception! no_echo = ' + convert(varchar(20), no_echo) to log;
			set no_echo = 0;
	end;

	--message 'trigger sdmc.wf_sdmc_bd::no_echo = ' + convert(varchar(20), no_echo) to log;
	if no_echo = 1 then
		return;
	end if;




	if (old_name.id_mat is not null) then
		call block_remote('stime', get_server_name(), 'mat');
		call delete_remote('stime', 'mat', 'id = ' + convert(varchar(20), old_name.id_mat));
		call unblock_remote('stime', get_server_name(), 'mat');
	end if;

	--message 'old_name.id_mat = ', old_name.id_mat to client;

	select sysname into remoteServer 
	from  guideventure v 
	join sdocs o on o.ventureId = v.ventureId and v.standalone = 0 and o.numdoc = old_name.numDoc and o.numext = old_name.numext;

	--message 'remoteServer = ', remoteServer to client;

	if remoteServer is not null and remoteServer != 'stime' then
		call block_remote(remoteServer, get_server_name(), 'mat');
		call delete_remote(remoteServer, 'mat', 'id = ' + convert(varchar(20), old_name.id_mat));
		call unblock_remote(remoteServer, get_server_name(), 'mat');
	end if;

end;


if exists (select 1 from systriggers where trigname = 'wf_sdmc_outcome_bi' and tname = 'sdmc') then 
	drop trigger sdmc.wf_sdmc_outcome_bi;
end if;

if exists (select 1 from systriggers where trigname = 'wf_sdmc_bi' and tname = 'sdmc') then 
	drop trigger sdmc.wf_sdmc_bi;
end if;

create 
	trigger wf_sdmc_bi before insert on 
sdmc
referencing new as new_name
for each row
begin
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_mat_nu varchar(20);
	declare v_currency_rate float;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_id_inv integer;
	declare v_id_source integer;
	declare v_id_dest integer;
	declare v_cost float;
	declare v_quant float;
	declare v_perList float;
	declare sync char(1);
	declare v_sysname varchar(32);
	declare v_venture_id integer;
	declare v_venture_anl integer;
	declare v_sysname_anl varchar(32);

	select id_jmat, s.id_voc_names, d.id_voc_names, n.ventureId, v.sysname
	into v_id_jmat, v_id_source, v_id_dest, v_venture_id, v_sysname
	from sdocs n 
		join sguidesource s on s.sourceid = n.sourid 
		join sguidesource d on d.sourceid = n.destid
		left join guideVenture v on n.ventureId = v.ventureId
	where n.numdoc = new_name.numdoc and n.numext = new_name.numext;

	if v_id_jmat is null then
		return;
	end if;

	set v_id_mat = get_nextid('mat');

	set v_id_currency = system_currency();
	call slave_currency_rate_stime(v_datev, v_currency_rate);
	call slave_select_stime(v_mat_nu, 'mat', 'max(nu)', 'id_jmat = ' + convert(varchar(20), v_id_jmat));
	
	set v_mat_nu = convert(varchar(20), convert(integer, isnull(v_mat_nu, 0)) + 1);

	select 
		id_inv
		, cost 
		, perList
	into 
		v_id_inv
		, v_cost 
		, v_perList
	from sguidenomenk 
	where nomnom = new_name.nomnom;


	set v_quant = new_name.quant;

	call get_venture_anl(v_venture_anl, v_sysname_anl);
	
	call wf_insert_mat (
		 v_sysname_anl
		,v_id_mat
		,v_Id_jmat
		,v_id_inv
		,v_mat_nu
		,v_quant 
		,v_cost
		,v_currency_rate
		,v_id_source
		,v_id_dest
		,v_perList
	);

	set new_name.id_mat = v_id_mat;

	if		isnull(v_venture_id, v_venture_anl) != v_venture_anl 
		and wf_dual_term(v_sysname, v_id_jmat) = 1
	then
		call wf_insert_mat (
			v_sysname
			,v_id_mat
			,v_Id_jmat
			,v_id_inv
			,v_mat_nu
			,v_quant 
			,v_cost
			,v_currency_rate
			,v_id_source
			,v_id_dest
			,v_perList
		);
	end if;
end;




----------------------------------------------------------------------
--------------                 sdocs          ------------------------
----------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_delete_sdocs' and tname = 'sdocs') then 
	drop trigger sdocs.wf_delete_sdocs;
end if;

create TRIGGER wf_delete_sdocs before delete on
sdocs
referencing old as old_name
for each row
begin
	declare remoteServer varchar(32);
	declare no_echo integer;

	set no_echo = 0;

  	begin
  		message '@stime_sdocs = ', @stime_sdocs to log;
		select @stime_sdocs into no_echo; 
	exception 
		when other then
			set no_echo = 0;
	end;

	if no_echo = 1 then
		return;
	end if;



	if (old_name.id_jmat is not null) then
		call block_remote('stime', get_server_name(), 'jmat');
		call block_remote('stime', get_server_name(), 'mat');
		call delete_remote('stime', 'jmat', 'id = ' + convert(varchar(20), old_name.id_jmat));
		call unblock_remote('stime', get_server_name(), 'jmat');
		call unblock_remote('stime', get_server_name(), 'mat');
	end if;

	select sysname into remoteServer 
	from  guideventure v 
	join orders o on o.ventureId = v.ventureId and v.standalone = 0 and o.numorder = old_name.numDoc;

--	message 'remoteServer = ', remoteServer to client;
	if remoteServer is not null and remoteServer != 'stime' then
		call block_remote(remoteServer, get_server_name(), 'jmat');
		call block_remote(remoteServer, get_server_name(), 'mat');
		call delete_remote(remoteServer, 'jmat', 'id = ' + convert(varchar(20), old_name.id_jmat));
		call unblock_remote(remoteServer, get_server_name(), 'jmat');
		call unblock_remote(remoteServer, get_server_name(), 'mat');
	end if;
end;


if exists (select 1 from systriggers where trigname = 'wf_set_numdoc' and tname = 'sdocs') then 
	drop trigger sdocs.wf_set_numdoc;
end if;


if exists (select 1 from systriggers where trigname = 'wf_insert_income' and tname = 'sdocs') then 
	drop trigger sdocs.wf_insert_income;
end if;


if exists (select 1 from sysprocedure where proc_name = 'change_mat_qty_dual') then
	drop procedure change_mat_qty_dual;
end if;

create
	procedure change_mat_qty_dual (
		  p_servername varchar(20)
		, p_id_mat integer
		, p_quant float
	) 
begin
	call call_remote ('stime', 'change_mat_qty', convert(varchar(20), p_id_mat) + ',' + convert(varchar(20), p_quant));
	if isnull(p_servername, 'stime') != 'stime' then
		call call_remote (p_servername, 'change_mat_qty', convert(varchar(20), p_id_mat) + ',' + convert(varchar(20), p_quant));
	end if;
end;


if exists (select 1 from sysprocedure where proc_name = 'wf_dual_insert_jmat') then
	drop procedure wf_dual_insert_jmat;
end if;

create
	procedure wf_dual_insert_jmat (
		  p_servername varchar(20)
		, p_id_guide_pmm integer
		, p_id_guide_anl integer
		, p_id_jmat integer
		, p_jmat_date date
		, p_jmat_nu integer
		, p_osn varchar(100)
		, p_id_currency integer
		, p_datev date
		, p_currency_rate float
		, p_id_s integer
		, p_id_d integer
		, p_id_jscet integer default 0
		, p_id_code integer default 0
	) 
begin

	call wf_insert_jmat (
		'stime'
		,p_id_guide_anl
		,p_id_jmat
		,p_jmat_date
		,p_jmat_nu
		,p_osn
		,p_id_currency
		,p_datev
		,p_currency_rate
		,p_id_s
		,p_id_d
		,p_id_jscet
		,p_id_code
	);

	if p_id_guide_pmm is not null then
		
		call wf_insert_jmat (
			 p_servername 
			,p_id_guide_pmm 
			,p_id_jmat
			,p_jmat_date
			,p_jmat_nu
			,p_osn
			,p_id_currency
			,p_datev
			,p_currency_rate
			,p_id_s
			,p_id_d
			,p_id_jscet
			,p_id_code
		);

	end if;
end;


if exists (select 1 from sysprocedure where proc_name = 'get_venture_anl') then
	drop procedure get_venture_anl;
end if;

create
	procedure get_venture_anl (
		  out p_venture_id integer
		, out p_sysname varchar(32)
	) 
begin
	--!todo
	set p_venture_id = 3;
	set p_sysname = 'stime';
end;


if exists (select 1 from sysprocedure where proc_name = 'wf_jmat_drop') then
	drop procedure wf_jmat_drop;
end if;

create
	procedure wf_jmat_drop (
		  in p_sysname varchar(32)
		, in p_id_jmat integer
	) 
begin
	if p_id_jmat is null or p_sysname is null then
		raiserror 17000 'Ошибка в параметрах wf_jmat_drop';
	end if;

	call delete_remote (p_sysname, 'jmat', 'id = ' + convert(varchar(20), p_id_Jmat));
end;


if exists (select 1 from sysprocedure where proc_name = 'wf_dual_term') then
	drop procedure wf_dual_term;
end if;

create
	function wf_dual_term (
		  in p_sysname varchar(32)
		, in p_id_jmat integer
	) returns integer
begin
	
	select count(*) 
	into wf_dual_term 
	from sdocs n
	join system s on 1=1
	where 
			n.id_jmat = p_id_jmat
		and n.xdate >= s.total_accounting_date; 
end;


if exists (select 1 from sysprocedure where proc_name = 'wf_jmat_distribute') then
	drop procedure wf_jmat_distribute;
end if;

create
	procedure wf_jmat_distribute (
		  in p_sysname varchar(32)
		, in p_id_jmat integer
		, in p_id_guide integer
		, in p_osn varchar(100) default null
	) 
begin
	declare v_tp1 integer;
	declare v_tp2 integer;
	declare v_tp3 integer;
	declare v_tp4 integer;
	declare f_inserted integer;
	declare v_osn varchar(100);
	declare v_mat_nu integer;
	declare v_id_currency integer;
	declare v_datev date;
	declare v_currency_rate float;

	if p_id_guide is null then
		return;
	end if;

	call qualify_guide(p_id_guide, v_tp1, v_tp2, v_tp3, v_tp4);

	set f_inserted = 0;
	set v_mat_nu = 1;

	for all_inc as iii sensitive cursor for
		select 
			  j.numdoc as r_numdoc
			, j.sourId as r_sourId, j.destId as r_destId
			, j.xDate as r_xdate
			, s.currency_iso as r_currency_iso
			, s.id_voc_names as r_id_s
			, d.id_voc_names as r_id_d
			, j.numext as r_numext
			, m.nomnom as r_nomnom
			, k.id_inv as r_id_inv
			, m.quant as r_quant
			, k.perlist as r_perlist
			, k.cost as r_cost
			, j.id_jmat as r_id_jmat
			, m.id_mat as r_id_mat
		from sdocs j 
			left join sdmc m on j.numdoc = m.numdoc and j.numext = m.numext 
			left join sguidenomenk k on k.nomnom = m.nomnom
			join sguidesource s on s.sourceid = j.sourid
			join sguidesource d on d.sourceid = j.destid
			left join guideventure v on v.ventureId = j.ventureId
		where j.id_jmat = p_id_jmat
		order by k.nomName
	do

		message '*>>>', r_numdoc, ' ', r_sourId, ' ', r_destId, ' ', r_currency_iso to client;
		if f_inserted = 0 then
			set f_inserted = 1;

			if p_osn is null then
				set v_osn = select_remote('stime', 'jmat', 'osn', 'id = ' + convert(varchar(20), isnull(p_id_jmat, -1)));
			else 
				set v_osn = p_osn;
			end if;

			set v_id_currency = select_remote('stime', 'jmat', 'id_curr', 'id = ' + convert(varchar(20), p_id_jmat));
			message '	v_id_currency = ', v_id_currency to client;
			set v_datev = select_remote('stime', 'jmat', 'datv', 'id = ' + convert(varchar(20), p_id_jmat));
			set v_currency_rate = select_remote('stime', 'jmat', 'curr', 'id = ' + convert(varchar(20), p_id_jmat));

			call wf_insert_jmat (
				 p_sysname
				,p_id_guide
				,p_id_jmat
				,r_xdate
				,r_numdoc
				,v_osn
				,v_id_currency
				,v_datev
				,v_currency_rate
				,r_id_s
				,r_id_d
			);
		end if;
		
		if r_id_mat is not null then
			call wf_insert_mat (
				 p_sysname
				,r_id_mat
				,p_id_jmat
				,r_id_inv
				,v_mat_nu
				,r_quant
				,r_cost
				,v_currency_rate
				,r_id_s
				,r_id_d
				,r_perList
			);
			set v_mat_nu = v_mat_nu + 1;
		end if;
	end for;

end;



if exists (select 1 from sysprocedure where proc_name = 'wf_id_guide_sdocs') then
	drop procedure wf_id_guide_sdocs;
end if;


if exists (select 1 from sysprocedure where proc_name = 'wf_jmat_id_guide') then
	drop procedure wf_jmat_id_guide;
end if;

create
	procedure wf_jmat_id_guide (
		  out o_id_guide_pmm integer
		, out o_id_guide_anl integer
		, out o_currency_iso varchar(10)
		, out o_id_currency  integer
		, out o_osn varchar(100)
		, p_venture_id integer -- через предприятие (ПМ или ММ). Может быть null
		, p_new_numext integer
		, p_new_sourId integer
		, p_new_destId integer
		, p_old_numext integer
	) 
begin

	declare v_ventureid_anl integer;


	if p_new_numext = 255 then
		-- по умолчанию - рублевый приход 
		set o_id_guide_anl = 1120;
		set o_currency_iso = 'RUR';
		set o_id_currency = 1;
		select 
			  c.id_guide
			, isnull(c.id_currency, ru.id_currency) 
			, s.currency_iso
		into 
			  o_id_guide_anl
			, o_id_currency
			, o_currency_iso
		from sguideSource s
		join GuideCurrency c on c.currency_iso = s.currency_iso
		join GuideCurrency ru on ru.currency_iso = 'RUR'
		where s.sourceId = p_new_sourId;
		set o_osn = 'Приход по накл. ';
	elseif (p_new_destId <= -1001 and p_new_sourId <= -1001) then
		set o_osn = 'Внутреннее перемещение ';
		set o_id_guide_anl = 1220;
	else
		set o_id_guide_anl = 1210;
		set o_osn = 'Расход по ';
	end if;

	set v_ventureid_anl = 3; -- todo! лучше взять из настроек system

	-- чтобы избежать повтора isnull(..)
	set p_venture_id = isnull(p_venture_id, v_ventureid_anl);

	if 
		-- если проводим заказ через аналитику
		p_venture_id = v_ventureid_anl
		-- или межсклад 
		or o_id_guide_anl = 1220
	then
		-- НЕ БУДЕМ СОЗДАВАТЬ НАКЛАДНУЮ В ПММ
		set o_id_guide_pmm = null;
	else
		if p_new_numext = 255 then
			-- приходуем в офиц. бухгалтерию накладные, как и в аналитику
			set o_id_guide_pmm = o_id_guide_anl;
		else
			-- чтобы не путались с автоформировательными накладными;
			set o_id_guide_pmm = 1217;
		end if;
	end if;
end;



if exists (select 1 from sysprocedure where proc_name = 'wf_jmat_shift_id') then
	drop procedure wf_jmat_shift_id;
end if;

create
	-- изменяет id для накладной так, что они становятся глобальными. включая id_mat
	-- изменению подлежат записи в stime и (если есть) в одной из офиц. бухгалтерий.
	function wf_jmat_shift_id (
		  in p_id_jmat integer
		, in p_ventureid integer default null
		, p_numdoc integer default null
		, p_numext integer default null
) returns integer
begin
	declare old_id_Jmat integer;
	declare v_id_mat integer;
	declare v_sysname varchar(20);
	declare v_sysname_anl varchar(20);

	set v_sysname_anl = 'stime'; --todo


	set wf_jmat_shift_id = null;

	if p_id_jmat is null then
		select id_jmat, v.sysname
		into old_id_jmat, v_sysname 
		from sdocs n
		left join guideventure v on v.ventureid = n.ventureid
		where numdoc = p_numdoc and numext = p_numext;
	else 
		set old_id_jmat = p_id_jmat;
		if p_ventureid is not null then
			select sysname into v_sysname from guideventure where ventureid = p_ventureid;
		end if;
	end if;	

	if old_id_jmat is null then
		raiserror 17000 'Error in procedure wf_jmat_shift_id. Text: old_id_jmat is null!';
		return;
	end if;

	message 'old_Id_jmat = ', old_id_jmat to client;

    for all_mat as m dynamic scroll cursor for
        select id_mat as r_id_mat from sdmc i
        join sdocs n on i.numdoc = n.numdoc and i.numext = n.numext
        where n.id_jmat = old_id_jmat
    do
        set v_id_mat = get_nextid('mat');
        call update_remote(v_sysname_anl, 'mat', 'id', v_id_mat, 'id = ' + convert(varchar(20), r_id_mat));
        if v_sysname != v_sysname_anl then
	        call update_remote(v_sysname, 'mat', 'id', v_id_mat, 'id = ' + convert(varchar(20), r_id_mat));
	    end if;
        update sdmc set id_mat = v_id_mat where id_mat = r_id_mat;
    end for;

    set wf_jmat_shift_id = get_nextid('jmat');
    message p_id_jmat to client;
    
    call update_remote(v_sysname_anl, 'jmat', 'id',  wf_jmat_shift_id, 'id = '+ convert(varchar(20), old_id_jmat));
	if v_sysname != v_sysname_anl then
		call update_remote(v_sysname, 'jmat', 'id',  wf_jmat_shift_id, 'id = '+ convert(varchar(20), old_id_jmat));
    end if;
    update sdocs set id_jmat = wf_jmat_shift_id where id_jmat = old_id_jmat;
end;

-- 
if exists (select 1 from sysprocedure where proc_name = 'wf_dual_distribute') then
	drop procedure wf_dual_distribute;
end if;

create
	procedure wf_dual_distribute (
		  p_numdoc         integer
		, p_numext         integer
		, p_sourid         integer
		, p_destid         integer
		, out o_id_jmat    integer
		, out o_venture_id integer
)
begin
--	declare v_id_jmat integer;
--	declare v_venture_id integer;
	declare v_id_mat integer;
	declare v_jmat_nu varchar(20);
	declare v_currency_rate float;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_id_source integer;
	declare v_id_dest integer;
	declare v_osn varchar(100);
	declare v_id_guide_jmat integer;
	declare v_currency_iso varchar(10);
	declare v_id_guide_anl integer;
	declare v_id_guide_pmm integer;
	declare v_sysname varchar(50);
	declare v_osn_type varchar(10);



	select ventureId into o_venture_id from orders where numorder = p_numdoc;

	if o_venture_id is null then
		select ventureId into o_venture_id from bayorders where numorder = p_numdoc;
		set v_osn_type = ' продаже ';
	else 
		set v_osn_type = ' заказу ';
	end if;

	if o_venture_id is not null then
		select sysname into v_sysname from guideVenture where ventureId = o_venture_id;
	else
		set v_osn_type = ' внутр. ';
	end if;

	call wf_jmat_id_guide (
		  v_id_guide_pmm, v_id_guide_anl, v_currency_iso, v_id_currency, v_osn
		, o_venture_id, p_numext, p_sourId, p_destId
		, null
	);

	set o_id_jmat = get_nextid('jmat');
	

	if isnull(v_currency_iso, 'RUR') = 'RUR' then
		set v_id_currency = system_currency();
	end if;

	call slave_currency_rate_stime(v_datev, v_currency_rate, null, v_id_currency);

	set v_jmat_nu = p_numdoc;
	select id_voc_names into v_id_source from sguidesource where sourceid = p_sourid;
	select id_voc_names into v_id_dest from sguidesource where sourceid = p_destid;
	set v_osn = v_osn + v_osn_type + convert(varchar(20), p_numdoc);
	if p_numext < 254 then
		set v_osn = v_osn + '/' + convert(varchar(20), p_numext);
	end if;
    
	call wf_dual_insert_jmat (
		 v_sysname
		,v_id_guide_pmm, v_id_guide_anl
		,o_id_jmat
		,now() --v_jmat_date
		,v_jmat_nu
		,v_osn
		,v_id_currency
		,v_datev
		,v_currency_rate
		,v_id_source
		,v_id_dest
	);

end;


-- триггер переименован wf_sdocs_bi
if exists (select 1 from systriggers where trigname = 'wf_sdocs_outcome_bi' and tname = 'sdocs') then 
	drop trigger sdocs.wf_sdocs_outcome_bi;
end if;

if exists (select 1 from systriggers where trigname = 'wf_sdocs_bi' and tname = 'sdocs') then 
	drop trigger sdocs.wf_sdocs_bi;
end if;

create 
	trigger wf_sdocs_bi before insert on 
sdocs
referencing new as new_name
for each row
--when (new_name.numext <= 254)
begin
	declare v_id_jmat integer;
	declare v_venture_id integer;
--	declare v_id_mat integer;
--	declare v_jmat_nu varchar(20);
--	declare v_currency_rate float;
--	declare v_datev date;
--	declare v_id_currency integer;
--	declare v_id_source integer;
--	declare v_id_dest integer;
--	declare v_osn varchar(100);
--	declare v_id_guide_jmat integer;
--	declare v_currency_iso varchar(10);
--	declare v_id_guide_anl integer;
--	declare v_id_guide_pmm integer;
--	declare v_sysname varchar(50);
--	declare v_osn_type varchar(10);


/*
	select ventureId into v_venture_id from orders where numorder = new_name.numdoc;

	if v_venture_id is null then
		select ventureId into v_venture_id from bayorders where numorder = new_name.numdoc;
		set v_osn_type = ' продаже ';
	else 
		set v_osn_type = ' заказу ';
	end if;

	if v_venture_id is not null then
		select sysname into v_sysname from guideVenture where ventureId = v_venture_id;
	else
		set v_osn_type = ' внутр. ';
	end if;

	call wf_jmat_id_guide (
		  v_id_guide_pmm, v_id_guide_anl, v_currency_iso, v_id_currency, v_osn
		, v_venture_id, new_name.numext, new_name.sourId, new_name.destId
		, null
	);

	set v_id_jmat = get_nextid('jmat');
	

	if isnull(v_currency_iso, 'RUR') = 'RUR' then
		set v_id_currency = system_currency();
	end if;

	call slave_currency_rate_stime(v_datev, v_currency_rate, null, v_id_currency);

	set v_jmat_nu = new_name.numdoc;
	select id_voc_names into v_id_source from sguidesource where sourceid = new_name.sourid;
	select id_voc_names into v_id_dest from sguidesource where sourceid = new_name.destid;
	set v_osn = v_osn + v_osn_type + convert(varchar(20), new_name.numdoc);
	if new_name.numext < 254 then
		set v_osn = v_osn + '/' + convert(varchar(20), new_name.numext);
	end if;
    
	call wf_dual_distribute (
		 v_sysname
		,v_id_guide_pmm, v_id_guide_anl
		,v_id_jmat
		,now() --v_jmat_date
		,v_jmat_nu
		,v_osn
		,v_id_currency
		,v_datev
		,v_currency_rate
		,v_id_source
		,v_id_dest
	);
*/
	call wf_dual_distribute (
		new_name.numdoc
		, new_name.numext
		, new_name.sourid
		, new_name.destid
		, v_id_jmat
		, v_venture_id 
	);
	set new_name.id_jmat = v_id_jmat;
	set new_name.ventureId = v_venture_id;

end;





if exists (select 1 from systriggers where trigname = 'wf_sdocs_outcome_bu' and tname = 'sdocs') then 
	drop trigger sdocs.wf_sdocs_outcome_bu;
end if;

if exists (select 1 from systriggers where trigname = 'wf_sdocs_bu' and tname = 'sdocs') then 
	drop trigger sdocs.wf_sdocs_bu;
end if;

create 
	trigger wf_sdocs_bu before update on 
sdocs
referencing new as new_name old as old_name
for each row
--when (old_name.numext = 254)
begin
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_jmat_nu varchar(20);
	declare v_currency_rate float;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_id_source integer;
	declare v_id_dest integer;
	declare v_osn varchar(100);

--	declare v_id_guide integer;
	declare v_tp1 integer;
	declare v_tp2 integer;
	declare v_tp3 integer;
	declare v_tp4 integer;
	declare v_currency_iso varchar(10);
	
	declare v_id_guide_jmat integer;
	declare v_id_guide_anl integer;
	declare v_id_guide_pmm integer;
	declare old_id_guide_pmm integer;
	declare v_venture_id integer;
	declare v_sysname varchar(50);
	declare v_venture_anl_Id integer;
	declare v_old_numext integer;
	declare old_id_guide_anl integer;
	declare f_distribute integer;
	declare v_total_accounting timestamp;
	declare chk_my_bug varchar(20);


	select venture_anl_id, total_accounting_date into v_venture_anl_id, v_total_accounting from system;

	set v_id_jmat = old_name.id_jmat;
	if update(ventureid) and new_name.ventureid is not null and new_name.ventureid != v_venture_anl_id and old_name.id_jmat is not null then
		select sysname into v_sysname from guideventure where ventureid = new_name.ventureid;
		set chk_my_bug = select_remote(v_sysname, 'jmat', 'id', 'id = ' + convert(varchar(20), old_name.id_jmat));
		if chk_my_bug is null then
			-- Сделать дополнительную проверку на глобальность id по позициям
			lopp:
			for all_mat as m dynamic scroll cursor for
			    select id_mat as r_id_mat from sdmc i
			    join sdocs n on i.numdoc = n.numdoc and i.numext = n.numext
			    where n.id_jmat = old_name.id_jmat
			do
				set chk_my_bug = select_remote(v_sysname, 'mat', 'id', 'id = ' + convert(varchar(20), r_id_mat));
				if chk_my_bug is not null then
					leave lopp;
				end if;
			end for;
		end if;
		if chk_my_bug is not null then
			-- Наткнулись на ситуация с неправильным (неглобальным) id для накладных.
			-- Тербуется перенести id в свободную область, включая id позиций номенклатурры.
			set v_id_jmat = wf_jmat_shift_id (
				  old_name.id_jmat 
				, old_name.ventureid
			);
		end if;
	end if;

	-- для тех накладных, которые относятся к переоду до интеграции 
	-- только для тех накладных которые уже имеет корресп. запись в Комтехе stime и в одной из официальных баз.
	if (update(ventureId) or update(sourId) or update (destId)) and v_id_jmat is not null  then
		-- при смене "от кого" и "кому" может произойти изменение типа накладной
		-- Поэтому нужно каждый раз проверять тип
		set v_old_numext = old_name.numext;
		call wf_jmat_id_guide (
			  v_id_guide_pmm, v_id_guide_anl, v_currency_iso, v_id_currency, v_osn
			, new_name.ventureid, new_name.numext, new_name.sourId, new_name.destId
			, v_old_numext
		);
--	end if;


--	if update (ventureId) then
		if		isnull(new_name.ventureId, v_venture_anl_id) != v_venture_anl_id
			and isnull(old_name.ventureId, v_venture_anl_id) != isnull(new_name.ventureId, v_venture_anl_id)
		then
			set f_distribute = 1;
		else 
			set f_distribute = 0;
		end if;
		message 'f_distribute = ',f_distribute to client;

		if v_id_jmat is not null then
			set old_id_guide_anl = select_remote('stime', 'jmat', 'id_guide', 'id = ' + convert(varchar(20), v_id_jmat));
			message 'old_id_guide_anl = ',old_id_guide_anl to client;
			message 'v_id_guide_anl = ',v_id_guide_anl to client;
			if old_id_guide_anl != v_id_guide_anl then
				call qualify_guide(
					  v_id_guide_anl
					, v_tp1
					, v_tp2
					, v_tp3
					, v_tp4
				);
	    
				call change_id_guide_remote (
					  'stime'
					, v_id_jmat
					, v_id_guide_anl
					, v_id_currency
					, v_tp1
					, v_tp2
					, v_tp3
					, v_tp4
				);
			end if;
		end if;

		if 		isnull(old_name.ventureId, v_venture_anl_id) != v_venture_anl_id 
			and old_name.xdate >= v_total_accounting
		then
			select sysname into v_sysname from guideventure where ventureid = old_name.ventureId;
		    -- исправить в базе старого предприятия если накладная меняет предприятие
		    if 
					isnull(new_name.ventureId, -old_name.ventureId) != old_name.ventureId 
				and old_name.ventureId != v_venture_anl_id
				and wf_dual_term(v_sysname, v_id_jmat) = 1
			then
				-- если предпирятие другое - удадляем накладную
				call wf_jmat_drop(v_sysname, v_id_jmat);
			else
				--set f_distribute = 0; -- предприятие осталось тем же, добавлять не нужно
				-- если тоже самое, то тогда проверяем, а может быть нужно поменять тип накладной
				set old_id_guide_pmm = select_remote(v_sysname, 'jmat', 'id_guide', 'id = ' + convert(varchar(20), v_id_jmat));
				if isnull(old_id_guide_pmm, -v_id_guide_pmm) != v_id_guide_pmm then

					call qualify_guide(
						  v_id_guide_anl
						, v_tp1
						, v_tp2
						, v_tp3
						, v_tp4
					);

					call change_id_guide_remote (
						  v_sysname
						, v_id_jmat
						, v_id_guide_pmm
						, v_id_currency
						, v_tp1
						, v_tp2
						, v_tp3
						, v_tp4
	    			);
    			end if;
			end if;
		end if;

		if		f_distribute = 1 
			and v_id_guide_pmm is not null 
			and old_name.xdate >= v_total_accounting
		then
			select sysname into v_sysname from guideventure where ventureid = new_name.ventureId;
			call wf_jmat_distribute(
					  v_sysname
					, v_id_jmat
					, v_id_guide_pmm
	    		);

	    	message 'after wf_jmat_distribute...' to client;

		end if;
	end if;
		
--		call qualify_guide(v_id_guide_pmm, v_tp1, v_tp2, v_tp3, v_tp4);

	if update(sourId) then
		select id_voc_names into v_id_source from sguidesource where sourceid = new_name.sourid;
		if v_Id_source is not null then
			call update_host('jmat', 'id_s', convert(varchar(20), v_id_source), 'id = ' + convert(varchar(20), v_id_jmat));
		end if;
	end if;
		
	if update(destId) then
		select id_voc_names into v_id_dest from sguidesource where sourceid = new_name.destid;
		if v_id_dest is not null then
			call update_host('jmat', 'id_d', convert(varchar(20), v_id_dest), 'id = ' + convert(varchar(20), v_id_jmat));
		end if;
	end if;

	if update(xDate) then
		call update_host('jmat', 'dat', '''''' + convert(varchar(20), new_name.xDate) + '''''', 'id = ' + convert(varchar(20), v_id_jmat));
	end if;

	--if update(note) then
		-- set v_osn = '[Prior: '+ new_name.note +']';
		-- пришлось отключить из-за ошибки при установки 
		-- признака предприятия в приходной накладной
		-- call update_remote ('stime', 'jmat', 'osn', '''' +v_osn + '''', 'id = ' + convert(varchar(20), v_id_jmat));
	--end if;
end;


---------------------------------------- inventory_order.sql ----------------------------------------

if exists (select 1 from sysprocedure where proc_name = 'inventory_order') then
	drop procedure inventory_order;
end if;

create
	procedure inventory_order (
		  p_inventory_date date default null
		, p_cost_preserve tinyint default null
		, p_sklad_id integer default null
	) 
begin

	declare v_id_inventar integer;
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_fields varchar(200);
	declare v_values varchar(2000);
	declare v_nu varchar(20);
	declare v_mat_nu integer;
	declare v_quant float;
	declare v_currency_rate real;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_osn varchar(100);

	declare sync char(1);
	declare f_update tinyint;
	declare f_insert_mat tinyint;
	declare f_insert_jmat tinyint;

	declare c_deleted varchar(20);

	set c_deleted = '-1000000';

	message 'inventory_order() started ...' to client;


	if p_inventory_date is null then
		set p_inventory_date = convert(date, now());
	end if;

	create table #saldo(nomnom varchar(20), id integer, debit float, kredit float);

	create table #itogo(nomnom varchar(20), id integer, debit float, kredit float);

	insert into #saldo (nomnom, id, debit)
    select nomnom, destid, sum(quant) from sdocs n
	join sdmc m on n.numdoc = m.numdoc and n.numext = m.numext
    where n.xDate < isnull(p_inventory_date, '30000101')
	group by m.nomnom, destid;
    
	insert into #saldo (nomnom, id, kredit)
    select nomnom, sourid, sum(quant) from sdocs n
	join sdmc m on n.numdoc = m.numdoc and n.numext = m.numext
    where n.xDate < isnull(p_inventory_date, '30000101')
	group by m.nomnom, sourid;
    

	insert into #itogo (nomnom, id, debit, kredit)
    select nomnom, id, sum(isnull(debit,0)), sum(isnull(kredit,0))
	from #saldo 
    group by nomnom, id;

--    begin
		call call_host('block_table', 'sync, ''prior'', ''jmat''');
		call call_host('block_table', 'sync, ''prior'', ''mat''');
--		call block_remote('stime', get_server_name(), 'jmat');
--		call block_remote('stime', get_server_name(), 'mat');

		


--		set v_currency_rate = system_currency_rate();
		set v_id_currency = system_currency();
		call slave_currency_rate_stime(v_datev, v_currency_rate);
		select id_voc_names into v_id_inventar from sguidesource where sourceName = 'Инвентаризация';


   	   	for sklad_cur as s dynamic scroll cursor for
			select 
				sourceid as r_sourceid
				, id_voc_names as r_id_sklad 
			from sguidesource
			where sourceid <= -1001
				and isnull(p_sklad_id, sourceid) = sourceid
		do

			
            -- глобальный для загловков накладных
            message p_inventory_date to client;
			set v_id_jmat = select_remote('stime', 'jmat', 'id'
				, 'convert(date, dat) = ''''' + convert(varchar(20), p_inventory_date) + ''''' and id_guide = 1023 and id_d = ' + convert(varchar(20), r_id_sklad));
			set f_update = 0;

			if v_id_jmat is null then
				set f_insert_jmat = 1;

				set v_id_jmat = get_nextid('jmat');
				call slave_select_stime(v_nu, 'jmat', 'max(nu)', 'id_guide = 1023');
				set v_nu = convert(varchar(20), convert(integer, isnull(v_nu, 0)) + 1);
			else
				if p_cost_preserve = 1 then
					-- будем исправлять только количество
					set f_update = 1;
					-- сохраняем в промежуточной колонке текущую цену, чтобы потом восстановить сумму по новому к-ву
					call update_remote('stime', 'mat', 'kol23', 'summa/kol1', 'id_jmat = ' + convert(varchar(20), v_id_jmat));
					-- помечаем, чтобы впоследствии удалить те позиции, которые не будут участвовать в новой инвентаризации
					call update_remote('stime', 'mat', 'kol1', c_deleted, 'id_jmat = ' + convert(varchar(20), v_id_jmat));
				else
					-- полностью обновить позиции инвентарицации (включая учетные цены)
					call delete_remote('stime', 'mat', 'id_jmat = ' + convert(varchar(20), v_id_jmat));
				end if;
			end if;


			if f_insert_jmat = 1 then
				call wf_insert_jmat (
					'stime'
					,'1023' --инветаризация
					,v_id_jmat
					,p_inventory_date
					,v_nu
					,v_osn
					,v_id_currency
					,v_datev
					,v_currency_rate
					,v_id_inventar
					,r_id_sklad
				);
			end if;

        	-- Добавляем предметы к накладной
        	if f_update != 1 then
	        	set v_mat_nu = 1;
	        else
	        	set v_mat_nu = select_remote('stime', 'mat', 'max(nu) + 1' , 'id_jmat = ' + convert(varchar(20), v_id_jmat));
	        end if;

			for nom_cur as n dynamic scroll cursor for
				select 
					i.nomnom as r_nomnom
					, n.id_inv as r_nomenklature_id
					, debit as r_debit, kredit as r_kredit 
					, cost as r_cost, perList as r_perList 
					, n.id_inv as r_id_inv
				from #itogo i
				join sguidenomenk n on n.nomnom = i.nomnom
	            where id = r_sourceid
					and isnull(p_sklad_id, id) = id
			do
				set v_quant = r_debit - r_kredit;

				if v_quant >= 0.01 then

					set f_insert_mat = 0;

					if f_update = 1 then
						set v_id_mat = select_remote('stime', 'mat', 'id'
							, 'id_jmat = ' + convert(varchar(20), v_id_jmat) + ' and id_inv = ' + convert(varchar(20), r_id_inv)) ;
						if v_id_mat is null then
							set f_insert_mat = 1;
						else
							call update_remote ('stime', 'mat', 'kol1', convert(varchar(20), v_quant/r_perlist)
								, 'id = ' + convert(varchar(20), v_id_mat));
							call update_remote ('stime', 'mat', 'summa', 'kol1 * kol23'
								, 'id = ' + convert(varchar(20), v_id_mat));
						end if;
					else
						set f_insert_mat = 1;
					end if;

					if f_insert_mat = 1 then
						message 'INSERT INTO MAT... ' to client;
						set v_id_mat = get_nextid('mat');
						call wf_insert_mat (
							'stime'
							,v_id_mat
							,v_Id_jmat
							,r_nomenklature_id
							,v_mat_nu
							,v_quant
							,r_cost
							,v_currency_rate
							,v_id_inventar
							,r_id_sklad
							,r_perList
						);
						--set v_id_mat = v_id_mat + 1;
						set v_mat_nu = v_mat_nu + 1;
					end if;

				end if;

			end for;
			if p_cost_preserve = 1 then
				call delete_remote('stime', 'mat', 'id_jmat = ' + convert(varchar(20), v_id_jmat) + ' and kol1 = ' + c_deleted);
			end if;
			set v_id_jmat = v_id_jmat + 1;
		end for;

		call unblock_remote('stime', get_server_name(), 'jmat');
		call unblock_remote('stime', get_server_name(), 'mat');
--		call call_host('unblock_table', 'sync, ''prior'', ''jmat''');
--		call call_host('unblock_table', 'sync, ''prior'', ''mat''');
--	exception 
--		when others then
--			set v_perList = v_perList;
--	end;

	drop table #saldo;
	drop table #itogo;
    
	message 'procedure inventory_order ended successful.' to client;
end;



if exists (select 1 from sysprocedure where proc_name = 'v_compensate_order') then
	drop procedure v_compensate_order;
end if;

create
-- Процедура инвентаризации по предприятию на дату
-- если первый параметр null - по всем придприятиям
-- если второй параметр null - на текущую дату
	procedure v_compensate_order (
		 p_venture_id integer default null
		, p_inventory_date date default null
		, p_total_start integer default 1
	) 
begin
	
	declare v_id_inventar integer;
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_fields varchar(200);
	declare v_values varchar(2000);
	declare v_nu varchar(20);
	declare v_mat_nu integer;
	declare v_quant float;
	declare v_currency_rate real;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_osn varchar(100);

	if p_inventory_date is null then
		set p_inventory_date = convert(date, now());
	end if;

	select id_voc_names into v_id_inventar from sguidesource where sourceName = 'Инвентаризация';
	set v_osn = 'Компенсация остатков до нулевого уровня (из-за некорректного учета раньше)';
		set v_id_jmat = get_nextid('jmat');

        -- глобальный для загловков накладных
		set v_id_mat = get_nextid('mat');
--		set v_currency_rate = system_currency_rate();
		set v_id_currency = system_currency();
		call slave_currency_rate_stime(v_datev, v_currency_rate);

   	for venture_cur as s dynamic scroll cursor for
		select 
			ventureid as r_ventureid
			, sysname as r_server
			, id_sklad as r_id_sklad
		from guideventure v
		where isnull(v.invCode, '' ) != '' and isnull(p_venture_id, v.ventureid) = v.ventureid
	do
		
			set v_nu = select_remote(r_server, 'jmat', 'max(nu)', 'id_guide = 1023');
			set v_nu = convert(varchar(20), convert(integer, isnull(v_nu, 0)) + 1);


			call wf_insert_jmat (
				r_server
				,'1023' --инветаризация
				,v_id_jmat
				,p_inventory_date
				,v_nu
				,v_osn
				,v_id_currency
				,v_datev
				,v_currency_rate
				,v_id_inventar
				,r_id_sklad
			);

        	-- Добавляем предметы к накладной
        	set v_mat_nu = 1;
			for nom_cur as n dynamic scroll cursor for
				select 
					  n.nomnom    as r_nomnom
					, n.id_inv    as r_nomenklature_id
					, n.cost      as r_cost
					, 1           as r_perlist
					, n.id_Inv    as r_id_inv
				from sguidenomenk n 
	            order by n.nomname
			do

        		call wf_calc_ost_inv_remote(r_server, v_quant, r_id_inv);

				if abs(v_quant) >= 0.01 then

--					select cost, perList into v_cost, v_perList from sguidenomenk where nomnom = r_nomnom;

					call wf_insert_mat (
						r_server
						,v_id_mat
						,v_Id_jmat
						,r_nomenklature_id
						,v_mat_nu
						,-v_quant
						,r_cost
						,v_currency_rate
						,v_id_inventar
						,r_id_sklad
						,r_perList
					);

					set v_id_mat = v_id_mat + 1;
					set v_mat_nu = v_mat_nu + 1;
				end if;

			end for;
			set v_id_jmat = v_id_jmat + 1;
	end for;
end;



if exists (select 1 from sysprocedure where proc_name = 'venture_inv_order') then
	drop procedure venture_inv_order;
end if;

if exists (select 1 from sysprocedure where proc_name = 'v_inventory_order') then
	drop procedure v_inventory_order;
end if;

create
-- Процедура инвентаризации по предприятию на дату
-- если первый параметр null - по всем придприятиям
-- если второй параметр null - на текущую дату
	procedure v_inventory_order (
		 p_venture_id integer default null
		, p_inventory_date date default null
		, p_total_start integer default 1
	) 
begin
	declare v_id_inventar integer;
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_fields varchar(200);
	declare v_values varchar(2000);
	declare v_nu varchar(20);
	declare v_mat_nu integer;
	declare v_quant float;
	declare v_currency_rate real;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_osn varchar(100);

	if p_inventory_date is null then
		set p_inventory_date = convert(date, now());
	end if;

	create table #saldo(nomnom varchar(20), id integer, debit float, kredit float);

	create table #itogo(nomnom varchar(20), id integer, debit float, kredit float);

	insert into #saldo (nomnom, id, debit, kredit)
	select r_nomnom, r_ventureid, sum(r_qty * r_kredit) as debit, 0
	from dummy
		join (
			select
				 quant as r_qty
				, m.nomnom as r_nomnom
				, if (n.sourid <= -1001 and n.destid <= -1001) then 
						0 
					else 
						if n.destid <= -1001 then 
							1
						else
							-1
   						endif
    			  endif 
	    			as r_kredit
    			, n.ventureid as r_ventureid
        	from sdocs n
    		join sdmc m on n.numdoc = m.numdoc and n.numext = m.numext 
--    		join sguidenomenk k on k.nomnom = m.nomnom
    		join sguidesource s on s.sourceId = n.sourId
    		join sguidesource d on d.sourceId = n.destId
    		join system sys on 1 = 1
    		join guideventure v on v.id_analytic = sys.id_analytic_default
    		left join orders o on o.numorder = n.numdoc
    		left join bayorders bo on bo.numorder = n.numdoc
			where
    			convert(date, n.xDate) <= isnull(p_inventory_date, convert(date, n.xDate))
    	) x on 1=1
	group by r_nomnom, r_ventureid;

	
	
		
	insert into #saldo (nomnom, id, debit, kredit)
    select m.nomnom, srcVentureId, 0, sum(m.quant) as kredit
			from sdmcventure m
			join sdocsventure n on m.sdv_id = n.id and n.cumulative_id is not null
--			join sguidenomenk k on k.nomnom = m.nomnom
			where n.nDate <= isnull(p_inventory_date, n.nDate)
			group by 
				m.nomnom, srcVentureId;

	insert into #saldo (nomnom, id, debit, kredit)
    select m.nomnom, dstVentureId, sum(m.quant) as kredit, 0
			from sdmcventure m
			join sdocsventure n on m.sdv_id = n.id and n.cumulative_id is not null
--			join sguidenomenk k on k.nomnom = m.nomnom
			where n.nDate <= isnull(p_inventory_date, n.nDate)
			group by 
				m.nomnom, dstVentureId;

	insert into #itogo (nomnom, id, debit, kredit)
	select s.nomnom, id, sum(debit), sum(kredit) 
	from #saldo s
	group by 
		s.nomnom, s.id;


	select id_voc_names into v_id_inventar from sguidesource where sourceName = 'Инвентаризация';
	set v_osn = 'Текущая инвентаризация';
		set v_id_jmat = get_nextid('jmat');

        -- глобальный для загловков накладных
		set v_id_mat = get_nextid('mat');
--		set v_currency_rate = system_currency_rate();
		set v_id_currency = system_currency();
		call slave_currency_rate_stime(v_datev, v_currency_rate);

   	for venture_cur as s dynamic scroll cursor for
		select 
			ventureid as r_ventureid
			, sysname as r_server
			, id_sklad as r_id_sklad
		from guideventure v
		where isnull(v.invCode, '' ) != '' and isnull(p_venture_id, v.ventureid) = v.ventureid
	do
		
			set v_nu = select_remote(r_server, 'jmat', 'max(nu)', 'id_guide = 1023');
			set v_nu = convert(varchar(20), convert(integer, isnull(v_nu, 0)) + 1);


			call wf_insert_jmat (
				r_server
				,'1023' --инветаризация
				,v_id_jmat
				,p_inventory_date
				,v_nu
				,v_osn
				,v_id_currency
				,v_datev
				,v_currency_rate
				,v_id_inventar
				,r_id_sklad
			);

        	-- Добавляем предметы к накладной
        	set v_mat_nu = 1;
			for nom_cur as n dynamic scroll cursor for
				select 
					i.nomnom as r_nomnom
					, n.id_inv as r_nomenklature_id
					, debit as r_debit
					, kredit as r_kredit 
					, n.cost as r_cost
					, n.perlist as r_perlist
				from #itogo i
				join sguidenomenk n on n.nomnom = i.nomnom
	            where i.id = r_ventureid
	            order by n.nomname
			do
				set v_quant = r_debit - r_kredit;

				if v_quant >= 0.01 then

--					select cost, perList into v_cost, v_perList from sguidenomenk where nomnom = r_nomnom;

					call wf_insert_mat (
						r_server
						,v_id_mat
						,v_Id_jmat
						,r_nomenklature_id
						,v_mat_nu
						,v_quant
						,r_cost
						,v_currency_rate
						,v_id_inventar
						,r_id_sklad
						,r_perList
					);

					set v_id_mat = v_id_mat + 1;
					set v_mat_nu = v_mat_nu + 1;
				end if;

			end for;
			set v_id_jmat = v_id_jmat + 1;
	end for;

	drop table #saldo;
	drop table #itogo;
end;




if exists (select 1 from sysprocedure where proc_name = 'venture_inv_qty') then
	drop function venture_inv_qty;
end if;

create
	-- возвращает остаток по позиции для заданного предприятия
	function venture_inv_qty (
		  p_nomnom varchar(20)
		, p_venture_id integer
		, p_inventory_date date default null
	) returns float
begin

    if p_nomnom is null or p_venture_id is null then
    	raiserror 17000 'Invalid parameter value';
    end if;

    if p_inventory_date is null then
    	set p_inventory_date = convert(date, now());
    end if;

	create table #saldo(id integer, debit float, kredit float);

	insert into #saldo (id, debit, kredit)
    select r_ventureid, sum(r_qty * r_kredit) as debit, 0
    from dummy
    	join (	
    		select
    				 quant/k.perlist as r_qty
    				, if (n.sourid <= -1001 and n.destid <= -1001) then 
    						0 
    					else 
    						if n.destid <= -1001 then 
    							1
    						else
    							-1
							endif
    					endif as 
					r_kredit
    				, if (n.sourid <= -1001 and n.destid <= -1001) then 
							null 
    					else 
    						if n.destid <= -1001 then 
    							isnull(n.ventureid, v.ventureid) 
    						else 
    							isnull(
    								isnull(
    									isnull(o.ventureid, bo.ventureid)
    									, if substring(isnull(o.invoice, bo.invoice), 1, 2) = '55' then 2 else 1 endif
    								), v.ventureid
    							) 
    						endif
    					endif 
					as r_ventureid 
        			from sdocs n
    				join sdmc m on n.numdoc = m.numdoc 
    						and n.numext = m.numext 
    				join sguidenomenk k on k.nomnom = m.nomnom
    			    join sguidesource s on s.sourceId = n.sourId
    				join sguidesource d on d.sourceId = n.destId
    				join system sys on 1 = 1
    				join guideventure v on v.id_analytic = sys.id_analytic_default
    				left join orders o on o.numorder = n.numdoc
    				left join bayorders bo on bo.numorder = n.numdoc
				where
    					m.nomnom = p_nomnom
					and convert(date, n.xDate) <= isnull(p_inventory_date, convert(date, n.xDate))
    	) x on 1=1
	group by r_ventureid;

	
	
		
	insert into #saldo (id, debit, kredit)
    select srcVentureId, 0, sum(m.quant / k.perlist) as kredit
			from sdmcventure m
			join sdocsventure n on m.sdv_id = n.id and n.cumulative_id is not null
			join sguidenomenk k on k.nomnom = m.nomnom
			where m.nomnom = p_nomnom
				and n.nDate <= isnull(p_inventory_date, n.nDate)
			group by srcVentureId;

	insert into #saldo (id, debit, kredit)
    select dstVentureId, sum(m.quant / k.perlist) as kredit, 0
			from sdmcventure m
			join sdocsventure n on m.sdv_id = n.id and n.cumulative_id is not null
			join sguidenomenk k on k.nomnom = m.nomnom
			where m.nomnom = p_nomnom
				and n.nDate <= isnull(p_inventory_date, n.nDate)
			group by dstVentureId;

	select sum(debit - kredit) 
	into venture_inv_qty
	from #saldo
	where id = p_venture_id;

	drop table #saldo;
end;


if exists (select 1 from sysprocedure where proc_name = 'inventory_tables_prep') then
	drop function inventory_tables_prep;
end if;

create
	procedure inventory_tables_prep ()
begin
	create table #saldo(nomnom varchar(20), id integer, debit float, kredit float);
	create table #itogo(nomnom varchar(20), id integer, debit float, kredit float);
    create table #nomnom(nomnom varchar(20) primary key);
end;	

if exists (select 1 from sysprocedure where proc_name = 'inventory_tables_clean') then
	drop function inventory_tables_clean;
end if;

create
	procedure inventory_tables_clean ()
begin
    drop table #nomnom;
	drop table #itogo;
	drop table #saldo;
end;	



if exists (select 1 from sysprocedure where proc_name = 'inventory_qty_rs') then
	drop function inventory_qty_rs;
end if;

create
	-- возвращает result set остаток по позиции если задана номенклатура p_nomnom
	-- Если номенклатура не задана, возвращаем 0, сохране
	function inventory_qty_rs (
		  p_nomnom varchar(20)
		, p_inventory_date date 
	) returns float
begin


	insert into #saldo (nomnom, id, debit)
    select nomnom, destid, sum(quant) from sdocs n
	join sdmc m on n.numdoc = m.numdoc and n.numext = m.numext
    where n.xDate < isnull(p_inventory_date, '30000101')
	    and m.nomnom = isnull(p_nomnom, m.nomnom)
	group by m.nomnom, destid;
    
	insert into #saldo (nomnom, id, kredit)
    select nomnom, sourid, sum(quant) from sdocs n
	join sdmc m on n.numdoc = m.numdoc and n.numext = m.numext
    where n.xDate < isnull(p_inventory_date, '30000101')
	    and m.nomnom = isnull(p_nomnom, m.nomnom)
	group by m.nomnom, sourid;
    

	insert into #itogo (nomnom, id, debit, kredit)
    select nomnom, id, sum(isnull(debit,0)), sum(isnull(kredit,0))
	from #saldo 
    group by nomnom, id;


end;


if exists (select 1 from sysprocedure where proc_name = 'inventory_qty') then
	drop function inventory_qty;
end if;

create
	-- возвращает остаток по позиции если задана номенклатура p_nomnom
	-- Если номенклатура не задана, возвращаем 0, сохране
	function inventory_qty (
		  p_nomnom varchar(20)
		, p_inventory_date date default null
		, p_sklad integer default null
	) returns float
begin

    if p_nomnom is null then
    	raiserror 17000 'Invalid parameter value';
    end if;

    if p_inventory_date is null then
    	set p_inventory_date = convert(date, now());
    end if;

    if p_perlist is null then
    	set p_perlist = 1;
    end if;

	call inventory_tables_prep();

	if p_nomnom is not null then
		insert into #nomnom (nomnom) select p_nomnom;
	else
		insert into #nomnom (nomnom) select nomnom from sguidenomenk;
	end if;

	insert into #saldo (nomnom, id, debit)
    select nomnom, destid, sum(quant) from sdocs n
	join sdmc m on n.numdoc = m.numdoc and n.numext = m.numext
	join #nomnom nm on nm.nomnom = m.nomnom
    where n.xDate < isnull(p_inventory_date, '30000101')
	group by m.nomnom, destid;
    
	insert into #saldo (nomnom, id, kredit)
    select nomnom, sourid, sum(quant) from sdocs n
	join sdmc m on n.numdoc = m.numdoc and n.numext = m.numext
	join #nomnom nm on nm.nomnom = m.nomnom
    where n.xDate < isnull(p_inventory_date, '30000101')
	group by m.nomnom, sourid;
    

	insert into #itogo (nomnom, id, debit, kredit)
    select nomnom, id, sum(isnull(debit,0)), sum(isnull(kredit,0))
	from #saldo 
    group by nomnom, id;

	message 'p_nomnom = ', p_nomnom to client;
	select (sum(isnull(debit, 0)) - sum(isnull(kredit, 0))) / p_perlist
	into inventory_qty
	from #saldo
		where id between isnull(p_sklad, -1002) and isnull(p_sklad, -1001)
	;

	call inventory_tables_clean();

end;

/*******************************************************************************************/

if exists (select '*' from sysprocedure where proc_name like 'wf_make_venture_income') then  
	drop function wf_make_venture_income;
end if;


create 
-- для накладной устанавливает признак, 
-- на какое предприятие был осуществлен приход
	function wf_make_venture_income (
	  p_numdoc varchar(20)
	, p_venture_id integer
	, p_numext integer default null
) returns integer
begin
	declare v_numdoc integer;
	declare v_numext integer;
	declare v_id_analytic_default integer;
	declare old_id_analytic integer;
	declare new_id_analytic integer;
	declare v_code_name varchar(30);
	declare v_id_jmat integer;
	declare v_activity_start date;
	declare v_xdate date;
	declare v_slash integer;
	declare v_ventureid integer;

	set v_slash = charindex('/', p_numdoc);
	if v_slash > 0 then
		set v_numdoc = convert (integer, substring (p_numdoc, 1, v_slash - 1));
		set v_numext = convert( integer, substring(p_numdoc, v_slash + 1));
	else 
		set v_numdoc = convert(integer, p_numdoc);
		set v_numext = p_numext;
/*
		if p_numext is null then
			set v_numext = 255;
		else
			set v_numext = p_numext;
		end if;
*/
	end if;

	set wf_make_venture_income = 1;
--	set v_numext = 255;
--	set v_numdoc = p_numdoc;

	select d.id_jmat, ov.id_analytic
		, v.id_analytic, s.id_analytic_default, v.activity_start, d.xdate
		, ov.ventureid
	into v_id_jmat, old_id_analytic, new_id_analytic, v_id_analytic_default, v_activity_start, v_xdate
		, v_ventureid
	from sdocs d
--	left join sdocsIncome i on i.numdoc = d.numdoc and i.numext = d.numext
	left join guideventure v on v.ventureId = p_venture_id
	left join guideventure ov on d.ventureId = ov.ventureId
	join system s on 1=1
	where d.numdoc = v_numdoc and d.numext = isnull(v_numext, numext);

	if v_activity_start > v_xdate then
		-- нельзя осуществить приход на предприятие до начала его работы
		set wf_make_venture_income = 0;
		return;
	end if;

	if p_venture_id != v_ventureid then
		update sdocs set ventureId = p_venture_id where numdoc = v_numdoc and numext = isnull(v_numext, numext);
	end if;

	if v_id_jmat is not null then
		call update_remote('stime', 'jmat', 'id_code', isnull(new_id_analytic, 0), 'id = ' + convert(varchar(20), v_id_jmat));
	else
--		set wf_make_venture_income = 0;
	end if;

	-- Приходуем накладную на то или иное предприятие в зависимости от 
	-- кода аналитики
	-- для этого в таблицу sDocsIncome добавляем/удаляем строку со
	-- ссылкой на предприятие
--	if new_id_analytic is null then
--		update sdocs set ventureId = null;
--		delete from sdocsincome where numdoc = v_numdoc and numext = v_numext;
--	else
--		if old_id_analytic is null then
--			insert into sdocsIncome (numdoc, numext, id_analytic, ventureId, id_jmat)
--			values (v_numdoc, v_numext, new_id_analytic, p_venture_id, v_id_jmat);
--	   	else 
--	   		update sdocsIncome set id_analytic = new_id_analytic 
--	   		where numdoc = v_numdoc and numext = v_numext;
--		end if;


--	end if;
	-- выставить признак того, что взаимозачеты необходимо пересчитать
	update sdocsventure dv set invalid = 1
	where v_xdate between dv.termFrom and dv.termTo
	and dv.cumulative_id is null;
end;
 

 


/* * * * * * * * * * * * * * * * * * * * * * * * * * *
 * Процедуры для получения себестоимости из комтеха
 * * * * * * * * * * * * * * * * * * * * * * * * * * */
if exists (select '*' from sysprocedure where proc_name like 'wf_get_comtex_cost') then  
	drop function wf_get_comtex_cost;
end if;



if exists (select '*' from sysprocedure where proc_name like 'wf_cost_bulk_change') then  
	drop function wf_cost_bulk_change;
end if;


create 
-- массовое обновление фактической цены для группы номенклатуры
	function wf_cost_bulk_change (
	p_klassid integer
	, p_cur_rate float default null
) returns integer
begin
	declare v_lvl integer;
	declare v_price_bulk_Id integer;
	declare v_comtex_cost float;
	declare v_timestamp datetime;
	declare v_cur_rate float;

	create table #tmp_klass(lvl integer, id integer);

	set v_lvl = 0;
	if p_klassid > 0 then
		insert into #tmp_klass (lvl, id) select 0, p_klassid;
	    
		branch: loop
			insert into #tmp_klass (lvl, id)
				select v_lvl + 1, k.klassId
				from sguideklass k
				join #tmp_klass t on t.id = k.parentKlassId and t.lvl = v_lvl;
	    
			if @@rowcount = 0 then
				leave branch;
			end if;
			set v_lvl = v_lvl + 1;
		end loop;
	else
		insert into #tmp_klass (lvl, id) 
		select 0, klassid
		from sguideklass
		where klassid != 0;
	end if;

	if p_cur_rate is not null then
		set v_cur_rate = p_cur_rate;
	else
		set v_cur_rate = system_currency_rate();
	end if;

	for v_table as b1 dynamic scroll cursor for
		select nomnom as r_nomnom, id_inv as r_id_inv
			, cost as r_prior_cost, perList as r_perlist
		from sguidenomenk n
		join #tmp_klass t on n.klassid = t.id
		where id_inv is not null
	do 
		call wf_calc_cost_stime(v_comtex_cost, r_id_inv);
		if v_comtex_cost > 0 then
			set v_comtex_cost = v_comtex_cost / v_cur_rate;
			if abs(round((v_comtex_cost - r_prior_cost), 2) ) > 0.01 then
				if v_price_bulk_Id is null then
					insert into sPriceBulkChange (guide_klass_id) values (p_klassid);
					set v_price_bulk_Id = @@identity;
				end if;
	    
				update sguidenomenk set cost = round(v_comtex_cost, 2) where nomnom = r_nomnom;
				-- триггером в этот момент добавляется запись в sPriceHistory
				select max(change_date) into v_timestamp from sPriceHistory where nomnom = r_nomnom;
				
				update sPriceHistory set bulk_id = v_price_bulk_id where change_date = v_timestamp and nomnom = r_nomnom;
	    
			end if;
		end if;
	end for;

	drop table #tmp_klass;

	return v_price_bulk_id;

end;

/* * * * * * * * * * * * * * * * * * * * * * * * * * *
 * Функция блокировки. (опять не работают)
 * * * * * * * * * * * * * * * * * * * * * * * * * * */

if exists (select '*' from sysprocedure where proc_name like 'bootstrap_blocking') then  
	drop procedure bootstrap_blocking;
end if;


create procedure bootstrap_blocking (
) 
begin

	
	call cre_block_var('blocks_inited');
	call cre_block_var('bulk_delete');
	call cre_block_var('supress_cum_update');
	call cre_block_var('supress_diary_update');

/*
	for v_table as b2 dynamic scroll cursor for
		select 'sdocs' as r_table union select 'sdmc' union select 'guidefirm' union select 'bayguidefirm'
	do 
		for v_server_name as a2 dynamic scroll cursor for
			select 
				srvname as r_server 
			from sys.sysservers s 
			join guideventure v on s.srvname = v.sysname
		do
			
			call cre_block_var(make_block_name(r_server, r_table));
		end for;
	end for;
*/
	
	for v_table as b1 dynamic scroll cursor for
		select 'jmat' as r_table union select 'mat' union select 'jscet' union select 'scet'
	do 
		for v_server_name as a1 dynamic scroll cursor for
			select 
				srvname as r_server 
			from sys.sysservers s 
			join guideventure v on s.srvname = v.sysname and v.standalone = 0
		do
			message 'call slave_cre_block_var_' + r_server + '(''' + make_block_name(get_server_name(), r_table) + ''')' to log;
			execute immediate 'call slave_cre_block_var_' + r_server + '(''' + make_block_name(get_server_name(), r_table) + ''')';

		end for;
	end for;



end;
	
	

if exists (select '*' from sysprocedure where proc_name like 'firstDayMonth') then  
	drop function firstDayMonth;
end if;


create function firstDayMonth (
	p_dt date
) returns date
begin
	declare v_total_day date;

	set firstDayMonth = 
		convert(date, ymd(year(p_dt), month(p_dt), 1));

	-- после перехода на полный учет сводные накладный перед переходом - не учитывать. 
	-- Их влияние заменено на компенсирующие и инвентаризационные накладные по предприятию.
	-- Период действия сводной накладной теперь - от дня следующего за днем перехода на тот.учет до конца месяца,
	-- иначе инвентаризационные накладные (которые датированы позже) "сожрут" взаимозачетую.
	select total_accounting_date into v_total_day from system;
	if p_dt >= v_total_day and firstDayMonth <= v_total_day then
		set firstDayMonth = v_total_day + 1;
	end if;
end;



if exists (select '*' from sysprocedure where proc_name like 'lastDayMonth') then  
	drop function lastDayMonth;
end if;


create function lastDayMonth (
	p_dt date
) returns date
begin
	declare v_total_day date;
	set lastDayMonth = 
			convert(date, ymd(year(p_dt), 1 + month(p_dt), 1) - 1);
	select total_accounting_date into v_total_day from system;
	if p_dt >= v_total_day and lastDayMonth <= v_total_day then
		set lastDayMonth = 
			convert(date, ymd(year(p_dt+1), 1 + month(p_dt+1), 1) - 1)
	end if;
end;


-------------------------------------------------------------------------
--------------       sDmcVenture triggers         ----------------------
-------------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_cumulative_del' and tname = 'sDmcVenture') then 
	drop trigger sDmcVenture.wf_cumulative_del;
end if;

create TRIGGER wf_cumulative_del before delete order 1 on
sDmcVenture
referencing old as old_name
for each row
when (exists (select 1 from sdocsventure n where n.id = old_name.sdv_id and n.cumulative_id is null))
begin
	declare v_cumulative_id integer;
	declare v_dmc_id integer;
	declare v_quant float;
	declare no_echo integer;

	select cumulative_id 
	into v_cumulative_id
	from sDocsVenture where id = old_name.sdv_id;

		-- удаляем позицию из сводной => нужно удалить все зачеты 
		-- из дневных накладных

--		execute immediate 'create variable @bulk_delete integer';
		set @bulk_delete = 1;
		select @bulk_delete into no_echo;

		delete from sdmcVenture diary
		from  
			  sDocsVenture diary_doc
		where diary.nomnom = old_name.nomnom
			and diary.sdv_id = diary_doc.id
			and diary_doc.cumulative_id = old_name.sdv_id
		;

		set @bulk_delete = 0;
--		execute immediate 'drop variable @bulk_delete';
end;

if exists (select 1 from systriggers where trigname = 'wf_diary_del' and tname = 'sDmcVenture') then 
	drop trigger sDmcVenture.wf_diary_del;
end if;

create TRIGGER wf_diary_del before delete order 2 on
sDmcVenture
referencing old as old_name
for each row
when (exists (select 1 from sdocsventure n where n.id = old_name.sdv_id and n.cumulative_id is not null))
begin
	declare v_cumulative_id integer;
	declare v_dmc_id integer;
	declare v_quant float;
	declare no_echo integer;

	select cumulative_id 
	into v_cumulative_id
	from sDocsVenture where id = old_name.sdv_id;

		-- скорректировать количество по позиции в сводной накладной

		
  	  	begin
			select @bulk_delete into no_echo; 
		exception 
			when other then
				set no_echo = 0;
		end;
	    
		if no_echo = 1 then
			return;
		end if;
		
		
		select 
			m.id
		into v_dmc_id
		from sDmcVenture m
		join sDocsVenture n on m.sdv_id = n.id 
			and n.id = v_cumulative_id
		where m.nomnom = old_name.nomnom
		;

		if v_dmc_id is not null then
			update sDmcVenture set quant = quant - old_name.quant 
			where id = v_dmc_id;
			select quant into v_quant from sDmcVenture where id = v_dmc_id;

			if round(v_quant, 3) < 0.001 then
				delete from sDmcVenture where id = v_dmc_id;
			end if;
			
		end if;
end;



if exists (select 1 from systriggers where trigname = 'wf_cumulative_upd' and tname = 'sDmcVenture') then 
	drop trigger sDmcVenture.wf_cumulative_upd;
end if;


create TRIGGER wf_cumulative_upd before update order 1 on
sDmcVenture
referencing old as old_name new as new_name
for each row
when (exists (select 1 from sdocsventure n where n.id = old_name.sdv_id and n.cumulative_id is null))
begin
	declare v_cumulative_id integer;
	declare v_dmc_id integer;
	declare v_quant float;
	declare v_ratio float;
	declare no_echo integer;
	declare v_cum_quant float;
	declare old_cum_costed float;
	declare old_cum_quant float;
	declare v_cumulative_total float;


  	  	begin
			select @supress_cum_update into no_echo; 
		exception 
			when other then
				set no_echo = 0;
		end;
	    
		if no_echo = 1 then
			return;
		end if;
		

		if update(costed) then
			-- пропорционально изменить сумму в дневных накладных
			-- так, чтобы сумма сводной накладной билась с суммой по дневными накладным
--			execute immediate 'create variable @supress_diary_update integer';
			set @supress_diary_update = 1;
			select @supress_diary_update into no_echo;

		    --message 'wf_cumulative_upd::old_name.costed = ', old_name.costed to client;
		    --message 'wf_cumulative_upd::new_name.costed = ', new_name.costed to client;
			if old_name.costed = 0 then
				update sdmcVenture diary set costed = new_name.costed
				from  
					sDocsVenture diary_doc
				where diary.nomnom = old_name.nomnom
					and diary.sdv_id = diary_doc.id
					and diary_doc.cumulative_id = old_name.sdv_id 
					;
			else 

				set v_ratio = new_name.quant * new_name.costed / (old_name.quant * old_name.costed);
		    
				update sdmcVenture diary set costed = costed * v_ratio
				from  
					 sDocsVenture diary_doc
				where diary.nomnom = old_name.nomnom
					and diary.sdv_id = diary_doc.id
					and diary_doc.cumulative_id = old_name.sdv_id 
					;
			end if;
	    
			set @supress_diary_update = 0;
--			execute immediate 'drop variable @supress_diary_update';
	    
		end if;
end;

if exists (select 1 from systriggers where trigname = 'wf_diary_upd' and tname = 'sDmcVenture') then 
	drop trigger sDmcVenture.wf_diary_upd;
end if;


create TRIGGER wf_diary_upd before update order 2 on
sDmcVenture
referencing old as old_name new as new_name
for each row
when (exists (select 1 from sdocsventure n where n.id = old_name.sdv_id and n.cumulative_id is not null))
begin
	declare v_cumulative_id integer;
	declare v_dmc_id integer;
	declare v_quant float;
	declare v_ratio float;
	declare no_echo integer;
	declare v_cum_quant float;
	declare old_cum_costed float;
	declare old_cum_quant float;
	declare v_cumulative_total float;



		begin
			select @supress_diary_update into no_echo; 
		exception 
			when other then
				set no_echo = 0;
		end;
	    
		if no_echo = 1 then
			return;
		end if;
		
--		    message 'wf_diary_upd::old_name.costed = ', old_name.costed to client;
--		    message 'wf_diary_upd::new_name.costed = ', new_name.costed to client;
	select cumulative_id 
	into v_cumulative_id
	from sDocsVenture where id = old_name.sdv_id;

		-- скорректировать количество по позиции в сводной накладной
		if update(quant) or update(costed) then

			select m.id, m.quant, m.costed
			into v_dmc_id, old_cum_quant, old_cum_costed
			from sDmcVenture m
			join sDocsVenture n on m.sdv_id = n.id and n.id = v_cumulative_id
			where m.nomnom = old_name.nomnom;

			if v_dmc_id is not null then

--				execute immediate 'create variable @supress_cum_update integer';
				set @supress_cum_update = 1;
				select @supress_cum_update into no_echo;

		    	set v_cum_quant = old_cum_quant - old_name.quant + new_name.quant;
		        set v_cumulative_total = (old_cum_quant * old_cum_costed) 
		        		- (old_name.quant * old_name.costed) 
		        		+ (new_name.quant * new_name.costed)
		        ;
	    
				update sDmcVenture 
					set costed = v_cumulative_total / v_cum_quant
					, quant = v_cum_quant
				where id = v_dmc_id;

				set @supress_cum_update = 0;
--				execute immediate 'drop variable @supress_cum_update';

			end if;
		end if;
end;



if exists (select 1 from systriggers where trigname = 'wf_cumulative_add' and tname = 'sDmcVenture') then 
	drop trigger sDmcVenture.wf_cumulative_add;
end if;

create TRIGGER wf_cumulative_add before insert order 1 on
sDmcVenture
referencing new as new_name
for each row
begin
	declare v_cumulative_id integer;
	declare v_dmc_id integer;
	declare no_echo integer;



	select cumulative_id 
	into v_cumulative_id
	from sDocsVenture where id = new_name.sdv_id;
	--message 'v_cumulative_id = ', v_cumulative_Id to client;


	if v_cumulative_id is not null then
		-- добавить (или проапдейтить) позицию в сводной накладной
		select 
			m.id
		into v_dmc_id
		from sDmcVenture m
		join sDocsVenture n on m.sdv_id = n.id 
			and n.id = v_cumulative_id
		where m.nomnom = new_name.nomnom
		;
		--message 'v_dmc_id = ', v_dmc_id to client;

		if v_dmc_id is null then
			insert into sDmcVenture (
				sdv_id, nomnom, quant, costed
			) values (
				v_cumulative_id
				, new_name.nomnom
				, new_name.quant
				, new_name.costed
			);
		else
--			execute immediate 'create variable @supress_cum_update integer';
			set @supress_cum_update = 1;
			select @supress_cum_update into no_echo;

			update sDmcVenture set quant = quant + new_name.quant 
			where id = v_dmc_id;

			set @supress_cum_update = 0;
--			execute immediate 'drop variable @supress_cum_update';
		end if;
			
	end if;
end;
-------------------------------------------------------------------------
--------------       end of sDmcVenture triggers         ---------------
-------------------------------------------------------------------------





-------------------------------------------------------------------------
--------------       sDocsVenture triggers         ----------------------
-------------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_cumulative_del' and tname = 'sDocsVenture') then 
	drop trigger sDocsVenture.wf_cumulative_del;
end if;

create TRIGGER wf_cumulative_del before delete order 1 on
sDocsVenture
referencing old as old_name
for each row
when (old_name.cumulative_id is null)
begin
end;



if exists (select 1 from systriggers where trigname = 'wf_cumulative_upd' and tname = 'sDocsVenture') then 
	drop trigger sDocsVenture.wf_cumulative_upd;
end if;

create TRIGGER wf_cumulative_upd before update order 1 on
sDocsVenture
referencing old as old_name new as new_name
for each row
when (old_name.cumulative_id is null)
begin
end;



if exists (select 1 from systriggers where trigname = 'wf_cumulative_add' and tname = 'sDocsVenture') then 
	drop trigger sDocsVenture.wf_cumulative_add;
end if;

create TRIGGER wf_cumulative_add before insert order 1 on
sDocsVenture
referencing new as new_name
for each row
when (new_name.cumulative_id is null)
begin
end;

if exists (select 1 from systriggers where trigname = 'wf_diary_del' and tname = 'sDocsVenture') then 
	drop trigger sDocsVenture.wf_diary_del;
end if;

create TRIGGER wf_diary_del before delete order 2 on
sDocsVenture
referencing old as old_name
for each row
when (old_name.cumulative_id is not null)
begin
	delete from sdmcVenture cum
	from sdmcVenture diary
	where cum.sdv_id = old_name.cumulative_id 
		and diary.sdv_id = old_name.id
		and cum.nomnom = diary.nomnom;
end;



if exists (select 1 from systriggers where trigname = 'wf_diary_upd' and tname = 'sDocsVenture') then 
	drop trigger sDocsVenture.wf_diary_upd;
end if;

create TRIGGER wf_diary_upd before update order 2 on
sDocsVenture
referencing old as old_name new as new_name
for each row
when (old_name.cumulative_id is not null)
begin
end;



if exists (select 1 from systriggers where trigname = 'wf_diary_add' and tname = 'sDocsVenture') then 
	drop trigger sDocsVenture.wf_diary_add;
end if;

create TRIGGER wf_diary_add before insert order 2 on
sDocsVenture
referencing new as new_name
for each row
when (new_name.cumulative_id is not null)
begin
	declare v_cumulative_id integer;
	-- предполагаем, что дневной взаимозачет вставляем с id = 0
	if new_name.cumulative_id = 0 then
		select id 
		into v_cumulative_id 
		from sDocsVenture
		where 
				termFrom = isnull(new_name.termFrom, firstDayMonth(new_name.nDate))
			and termTo   = isnull(new_name.termTo, lastDayMonth(new_name.nDate))
			and srcVentureId = new_name.srcVentureId
			and dstVentureId = new_name.dstVentureId
			and cumulative_id is null;

		if v_cumulative_Id is null then
			insert into sdocsventure (
				termFrom
				, termTo
				, srcVentureId
				, dstVentureId
				, cumulative_id
				, nDate
				, procent
			) values (
				  isnull(new_name.termFrom, firstDayMonth(new_name.nDate))
				, isnull(new_name.termTo, lastDayMonth(new_name.nDate))
				, new_name.srcVentureId
				, new_name.dstVentureId
				, null
				, firstDayMonth(new_name.nDate)
				, new_name.procent
			);
			set v_cumulative_id = @@identity;
		end if;
		set new_name.cumulative_id = v_cumulative_id;
	end if;

	if new_name.termFrom is null then
		set new_name.termFrom = firstDayMonth(new_name.nDate);
	end if;
	if new_name.termTo is null then
		set new_name.termTo = lastDayMonth(new_name.nDate);
	end if;
end;

-------------------------------------------------------------------------
--------------       end of sDocsVenture triggers         ---------------
-------------------------------------------------------------------------

if exists (select '*' from sysprocedure where proc_name like 'ivo_validate') then  
	drop procedure ivo_validate;
end if;


create 
	-- Проверка признака были ли исправления задним числом в накладных
	-- Если да, то вызывается пересчет накладной
	procedure ivo_validate (
	  p_procentOver float default null
) 
begin
	declare v_invalidate integer;
	declare v_term_min date;
	declare v_term_max date;

	set v_invalidate = 0;
	for ivo_c as ivo dynamic scroll cursor for
		select 
			id_jmat as r_id_jmat 
			, n.id as r_ivo_id
			, d.id_analytic as r_id_analytic
			, termFrom as r_term_start
			, termTo as r_term_end
			, nDate as r_nDate
			, isnull(invalid, 0) as r_invalid
			, n.procent as r_ivo_procent
		from sDocsVenture n
		join guideVenture s on s.ventureId = n.srcVentureId
		join guideVenture d on d.ventureId = n.dstVentureId
--		where isnull(n.invalid, 0) = 1
		order by n.ndate
	do
		if r_invalid = 1 then
			set v_invalidate = 1;
			update sdocsventure set invalid = 0 where id = r_ivo_id;
		end if;
		if v_invalidate = 1 then
			call ivo_comtex_delete(r_ivo_id);
			delete from sdocsventure where cumulative_Id = r_ivo_id;
			--delete from sdmcventure where sdv_id = r_ivo_id;
			if r_term_start <= isnull(v_term_min, '20000101') then
				set v_term_min = r_term_start;
			end if;
			if r_term_end >= isnull(v_term_max, '21000101') then
				set v_term_max = r_term_end;
			end if;
		end if;
	end for;
	call ivo_generate(
		p_procentOver
		, v_term_min
		, v_term_max
	);
end;


if exists (select '*' from sysprocedure where proc_name like 'ivo_comtex_delete') then  
	drop procedure ivo_comtex_delete;
end if;


create 
-- Удаляет из базы Комтеха информацию о взаимозачетной накладной
procedure ivo_comtex_delete (
	 p_ivo_id integer
) 
begin

	for ivo_c as ivo dynamic scroll cursor for
		select 
			id_jmat as r_id_jmat 
			, d.id_analytic as r_id_analytic
			, termFrom as r_term_start
			, termTo as r_term_end
			, nDate as r_nDate
			, s.sysname as src_sysname
			, d.sysname as dst_sysname
		from sDocsVenture n
		join guideVenture s on s.ventureId = n.srcVentureId
		join guideVenture d on d.ventureId = n.dstVentureId
		where id = p_ivo_id
	do
		if r_id_jmat is not null then
			call block_remote(src_sysname, get_server_name(), 'jmat');
			call block_remote(src_sysname, get_server_name(), 'mat');

			call delete_remote(src_sysname, 'jmat', 'id = '+ convert(varchar(20), r_id_jmat));

			call unblock_remote(src_sysname, get_server_name(), 'mat');
			call unblock_remote(src_sysname, get_server_name(), 'jmat');

			call block_remote(dst_sysname, get_server_name(), 'jmat');
			call block_remote(dst_sysname, get_server_name(), 'mat');

			call delete_remote(dst_sysname, 'jmat', 'id = '+ convert(varchar(20), r_id_jmat));

			call unblock_remote(dst_sysname, get_server_name(), 'mat');
			call unblock_remote(dst_sysname, get_server_name(), 'jmat');
			update sdocsVenture set id_jmat = null where id = p_ivo_id;
		end if;
	end for;
	
end;




if exists (select '*' from sysprocedure where proc_name like 'ivo_to_comtex') then  
	drop procedure ivo_to_comtex;
end if;


create 
-- Перевести информацию о взаимозачете из базы приора в базу Комтеха
-- требует переделки.
procedure ivo_to_comtex (
	 p_ivo_id integer
) 
begin
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_jmat_nu varchar(20);
	declare v_mat_nu integer;
	declare v_currency_rate float;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_id_source integer;
	declare v_id_dest integer;
	declare v_osn varchar(100);
	declare v_id_guide_jmat integer;
	declare v_folder_id integer;
	declare char_id_jmat varchar(20);

	declare v_src_id_sklad  integer;
	declare v_src_id_guide  integer;
	declare v_src_osn       varchar(100);
	declare v_dst_id_sklad  integer;
	declare v_dst_id_guide  integer;
	declare v_dst_osn       varchar(100);


	declare v_note varchar(50);



	for ivo_c as ivo dynamic scroll cursor for


		select 
			id_jmat as r_id_jmat 
			, d.id_analytic as r_id_analytic
			, termFrom as r_term_start
			, termTo as r_term_end
			, nDate as r_nDate
			, s.rusAbbrev as r_srcAbbrev
			, d.rusAbbrev as r_dstAbbrev
			, s.sysname as r_src_server
			, d.sysname as r_dst_server
			, n.srcVentureId as r_src_venture_id
			, n.dstVentureId as r_dst_venture_id
		from sDocsVenture n
		join guideVenture s on s.ventureId = n.srcVentureId
		join guideVenture d on d.ventureId = n.dstVentureId
		where id = p_ivo_id
	do

		call ivo_recognize_jmats (
			  r_src_venture_id 
			, r_dst_venture_id 
			, v_src_id_sklad 
			, v_src_id_guide 
			, v_src_osn 
			, v_dst_id_sklad 
			, v_dst_id_guide 
			, v_dst_osn 
		);

		set v_id_currency = system_currency();
		call slave_currency_rate_stime(v_datev, v_currency_rate);
		set v_jmat_nu = convert(varchar(20), p_ivo_id);


		message 'dst_sklad =', v_dst_id_sklad, ', src_sklad =', v_src_id_sklad to client;
		if v_src_id_sklad is not null then
			if v_dst_id_sklad is not null then
				select id_voc_names into v_id_dest from sguidesource where sourceName like '%взаимозачет%' and sourceid < 0;
			else
				select id_voc_names into v_id_dest from sguidesource where sourceName like '%инвентаризация%' and sourceid < 0;
			end if;
			message 'v_id_dest =', v_id_dest to client;

			if r_id_jmat is not null then
				-- проверим, не удалена ли она в Комехе?
				set char_id_jmat = select_remote(r_src_server, 'jmat', 'id', 'id = ' + convert(varchar(20), r_id_jmat));
				if char_id_jmat is not null then
					-- удалить деталировку
					-- добавление 26.11.06: вместе с самой накладной 
					-- деталировка удаляется по foreign constraint
					call delete_remote(r_src_server, 'jmat', 'id = '+ char_id_jmat);
				end if;
			else 
				set r_id_jmat = get_nextid('jmat');
				set char_id_jmat = null;
				update sDocsVenture d set 
					id_jmat = r_id_jmat
				where 
					id = p_ivo_id;
			end if;
			message 'r_id_jmat =', r_id_jmat to client;

			call block_remote(r_src_server, get_server_name(), 'jmat');
			call wf_insert_jmat (
				 r_src_server
				,v_src_id_guide
				,r_id_jmat
				,r_nDate
				,v_jmat_nu
				,v_src_osn
				,v_id_currency
				,v_datev
				,v_currency_rate
				,v_src_id_sklad
				,v_id_dest
				,0 -- id_jscet
				,0 -- r_id_analytic
			);
			call unblock_remote(r_src_server, get_server_name(), 'jmat');


		end if;

		if v_dst_id_sklad is not null then
			if v_src_id_sklad is not null then
				select id_voc_names into v_id_source from sguidesource where sourceName like '%взаимозачет%' and sourceid < 0;
			else
				select id_voc_names into v_id_source from sguidesource where sourceName like '%инвентаризация%' and sourceid > 0;
			end if;
			message 'v_id_source =', v_id_source to client;

			if r_id_jmat is not null then
				-- проверим, не удалена ли она в Комехе?
				set char_id_jmat = select_remote(r_dst_server, 'jmat', 'id', 'id = ' + convert(varchar(20), r_id_jmat));
				if char_id_jmat is not null then
					-- удалить деталировку
					call delete_remote(r_dst_server, 'jmat', 'id = '+ char_id_jmat);
				end if;
			else 
				set r_id_jmat = get_nextid('jmat');
				set char_id_jmat = null;
				update sDocsVenture d set 
					id_jmat = r_id_jmat
				where 
					id = p_ivo_id;
			end if;

			message 'r_id_jmat =', r_id_jmat to client;
			call block_remote(r_dst_server, get_server_name(), 'jmat');
			call wf_insert_jmat (
				 r_dst_server
				,v_dst_id_guide
				,r_id_jmat
				,r_nDate
				,v_jmat_nu
				,v_dst_osn
				,v_id_currency
				,v_datev
				,v_currency_rate
				,v_id_source
				,v_dst_id_sklad
				,0 -- id_jscet
				,0 -- r_id_analytic
			);
			call unblock_remote(r_dst_server, get_server_name(), 'jmat');
		end if;

		
		set v_mat_nu = 1;
		-- добавить деталировку по сводной накладной
		for anomnom as c_nomnom dynamic scroll cursor for
			select 
				  quant as r_qty
				, m.nomnom as r_nomnom
				, perlist as r_perList
				, m.costed as r_cost
				, k.id_inv as r_id_inv
			from sdocsventure d
			join sdmcventure m on m.sdv_id = d.id
			join sguidenomenk k on k.nomnom = m.nomnom
			where 
				d.id_jmat = r_id_jmat
		do
			
			set v_id_mat = null;
			--message r_nomnom to client;

			if v_src_id_sklad is not null then
				call block_remote(r_src_server, get_server_name(), 'mat');
				set v_id_mat = wf_insert_mat (
					r_src_server
					,v_id_mat
					,r_id_jmat
					,r_id_inv
					,v_mat_nu
					,r_qty 
					,r_cost
					,v_currency_rate
					,v_src_id_sklad
					,v_id_dest
					,r_perList
				);
				call unblock_remote(r_src_server, get_server_name(), 'mat');
			end if;
	    
			if v_dst_id_sklad is not null then
				call block_remote(r_dst_server, get_server_name(), 'mat');
				set v_id_mat = wf_insert_mat (
					r_dst_server
					,v_id_mat
					,r_id_jmat
					,r_id_inv
					,v_mat_nu
					,r_qty 
					,r_cost
					,v_currency_rate
					,v_id_source
					,v_dst_id_sklad
					,r_perList
				);
				call unblock_remote(r_dst_server, get_server_name(), 'mat');
			end if;
	    
	        update sdmcventure m set m.id_mat = v_id_mat 
	        where   r_id_jmat = m.sdv_id
				and m.nomnom = r_nomnom
			;

			set v_mat_nu = v_mat_nu + 1;

		end for;
	end for;
end;



if exists (select '*' from sysprocedure where proc_name like 'wf_put_ivo_nomnom') then  
	drop function wf_put_ivo_nomnom;
end if;


create 
    -- загрузка информации о взаимозачету по позиции номенклатуры на дату p_target_date
    -- между предприятиями p_srcVentureId и p_dstVentureId
    -- Функция ищет дневную накладную и если находит использует ее
    -- Если нет, то создает новую дневную накладную
	function wf_put_ivo_nomnom (
	  p_target_date date
	, p_nomnom varchar(50)
	, p_qty    float
	, p_procent float
	, p_srcVentureid integer
	, p_dstVentureId integer
	, p_term_start date    default null
	, p_term_end date      default null
) returns integer
begin
	declare v_ndate date;
	declare v_nomnom varchar(20);
	declare v_costed float;
	declare v_perList float;
	declare v_procent float;
	declare chk_forward_income_qty float;
	declare v_comtex_cost float;
	declare v_id_inv integer;

	select d.id, m.nomnom, isnull(m.quant, 0), procent
	into wf_put_ivo_nomnom, v_nomnom, chk_forward_income_qty, v_procent
	from sdocsventure d
	left join sdmcventure m on m.sdv_id = d.id and m.nomnom = p_nomnom
	where d.nDate = p_target_date 
		and srcVentureId = p_srcVentureId
		and dstVentureId = p_dstVentureId
 		and d.cumulative_id is not null
	;

	if wf_put_ivo_nomnom is null then
		insert into sDocsVenture (nDate, srcVentureId, dstVentureId, procent, termFrom, termTo)
		values (p_target_date, p_srcVentureId, p_dstVentureId, isnull(p_procent, v_procent), p_term_start, p_term_end);
		set wf_put_ivo_nomnom = @@identity;
	end if;

	select cena1, perList into v_costed, v_perList from sguidenomenk where nomnom = p_nomnom;


	if v_nomnom is null then
		insert into sDmcVenture(sdv_id, nomnom, quant, costed)
		select wf_put_ivo_nomnom, p_nomnom, p_qty * v_perList, v_costed * (1 + p_procent / 100);
	else 

		update sDmcVenture set quant = quant + p_qty * v_perList
		where sdv_id = wf_put_ivo_nomnom and nomnom = p_nomnom;
	end if;
/*
	--  Взять не текущую себестоимость во взаимозачете, а цену на дату взаимозачета.
	-- Не проходит сейчас, потому что нужно в этом случае выправлять ситуацию в 2004 и 2005 годах
	-- а это делать не очень целесообразно
	select id_inv, perList into v_id_inv, v_perList from sguidenomenk where nomnom = p_nomnom;
	call wf_cost_date_stime(v_comtex_cost, v_id_inv, p_target_date);

	set v_comtex_cost = v_comtex_cost /30;

	if v_nomnom is null then
		insert into sDmcVenture(sdv_id, nomnom, quant, costed)
		select wf_put_ivo_nomnom, p_nomnom, p_qty * v_perList, v_comtex_cost * (1 + p_procent / 100);
	else 

		update sDmcVenture set quant = quant + p_qty * v_perList, costed = v_comtex_cost * (1 + p_procent / 100)
		where sdv_id = wf_put_ivo_nomnom and nomnom = p_nomnom;
	end if;
*/	
end;


if exists (select '*' from sysprocedure where proc_name like 'ivo_recognize_jmats') then  
	drop procedure ivo_recognize_jmats;
end if;

create 
	-- При переносе взаимоозачетов в Комтех требуются получить данные по складу и типу накладной 
	-- в зависимости от того, какие предприятия задействованы в нем.
	procedure ivo_recognize_jmats (
		  p_src_venture_id integer
		, p_dst_venture_id integer
		, out o_src_id_sklad integer
		, out o_src_id_guide integer
		, out o_src_osn varchar(100)
		, out o_dst_id_sklad integer
		, out o_dst_id_guide integer
		, out o_dst_osn varchar(100)
)

begin
	declare v_analytic_id integer;
	declare v_src_abbrev varchar(20);
	declare v_dst_abbrev varchar(20);

	select rusAbbrev into v_src_abbrev from GuideVenture where ventureId = p_src_venture_id;

	select rusAbbrev into v_dst_abbrev from GuideVenture where ventureId = p_dst_venture_id;


	select id_sklad into o_src_id_sklad from guideventure where ventureId = p_src_venture_id;
	set o_src_id_guide = get_id_guide_by_key('outcome', 1); -- валютные - чтобы не путались с "нормальными" по автоформированию
	set o_src_osn = 'Расход по взаимозачету на ' + v_dst_abbrev;

	select id_sklad into o_dst_id_sklad from guideventure where ventureId = p_dst_venture_id;
	set o_dst_id_guide = get_id_guide_by_key('income', 1);
	set o_dst_osn = 'Приход по взаимозачету от ' + v_src_abbrev;

	set v_analytic_id = 3; -- todo
	if p_src_venture_id = v_analytic_id then
		set o_src_id_sklad = null;
		set o_src_id_guide = null;
		set o_dst_osn = 'Приход по инвентаризации';
	elseif p_dst_venture_id = v_analytic_id then
		set o_dst_id_sklad = null;
		set o_dst_id_guide = null;
		set o_src_osn = 'Расход на внутр. цели';
	end if;
end;



if exists (select '*' from sysprocedure where proc_name like 'ivo_generate_nomnom') then  
	drop procedure ivo_generate_nomnom;
end if;

create 
	-- автоматическое формирование взаимозачета между предприятиями по одной номенклатуре
	-- 
	procedure ivo_generate_nomnom (
	  p_nomnom varchar(50)
	, p_procentOver float
	, p_defaultVentureId integer
	, p_term_start date    default null
	, p_term_end date      default null
)

begin
	declare total_rest float;
	declare rest1 float;
	declare rest2 float;
--	declare v_defaultVentureId integer;
	declare v integer;
	declare cnt integer;
	declare vo_summa float;
	declare ivo_id integer;

	message 'Generate ivo for ', p_nomnom to client;

		update #vntRest set rest = 0.00;
		
		nomnom_loop:
		for cur_history as his sensitive cursor for
			select 
				if destId <= -1001 then 2 else 3 endif 
    				as sec_sort
				, convert(date, xDate) as r_nDate
				, n.sourid 
				, n.destid 
				, quant/k.perlist as r_qty
				, 	if (n.sourid <= -1001 and n.destid <= -1001) then 
						0 
					else 
						if n.destid <= -1001 then 
							1
						else
							-1
						endif
					endif as 
				r_activeOper
				, n.ventureid as r_ventureid 
				, 0 as r_destVentureId
				, convert(varchar(20), n.numdoc) + '/' + convert(varchar(20),n.numext) as r_numdoc
			from sdocs n
				join sdmc m on n.numdoc = m.numdoc 
						and n.numext = m.numext 
				join sguidenomenk k on k.nomnom = m.nomnom
			    join sguidesource s on s.sourceId = n.sourId
				join sguidesource d on d.sourceId = n.destId
				join system sys on 1 = 1
				join guideventure v on v.id_analytic = sys.id_analytic_default
				left join orders o on o.numorder = n.numdoc
				left join bayorders bo on bo.numorder = n.numdoc
			where
					m.nomnom = p_nomnom
				and convert(date, n.xDate) <= isnull(p_term_end, convert(date, n.xDate))
						union
			select 
				  1 as sec_sort 
				, n.nDate as r_nDate
				, null as sourId, null as destId
				, m.quant / k.perlist as r_qty
				, 0 as r_activeOper
				, srcVentureId as r_ventureId
				, dstVentureId as r_destVentureId
				, convert(varchar(20), n.id) as r_numdoc
			from sdmcventure m
			join sdocsventure n on m.sdv_id = n.id and n.cumulative_id is not null
			join sguidenomenk k on k.nomnom = m.nomnom
			where m.nomnom = p_nomnom
				and n.nDate <= isnull(p_term_end, n.nDate)
			order by 2, 1
		do
--			if r_nDate > isnull(p_term_end, r_nDate) then
--				leave nomnom_loop;
--			end if;

			if r_destVentureId > 0 then
				update #vntRest set rest = rest + r_qty where ventureId = r_destVentureId;
				update #vntRest set rest = rest - r_qty where ventureId = r_ventureId;
			else
				if not exists (select 1 from #vntRest where ventureId = r_ventureId) then
					set r_ventureId = p_defaultVentureId;
				end if; 
				update #vntRest set rest = rest + r_qty * r_activeOper where ventureId = r_ventureId;
			end if;
			
			if exists (select 1 from #vntRest where rest < 0)
				and r_nDate >= isnull(p_term_start, r_nDate)
--				and r_nDate <= isnull(p_term_end, r_nDate)
			then
				select sum(rest) into total_rest from #vntRest;
				if abs(round(total_rest, 3)) >= 0 then 
					-- найти где образовался минус, чтобы его компенсировать за счет того, у кого плюс
					compensate:
					for dst_vent as dv sensitive cursor for
						select rest as dv_rest, ventureId as dv_dstVentureId 
						from #vntRest where round(rest, 3) < 0
					do
						--message 'r_ndate = ', r_ndate, '    dv_rest = ', round(dv_rest, 2), '     total_rest =', round(total_rest, 2), '     r_ventureId = ', r_ventureId to client;
						set vo_summa = truncnum( abs(dv_rest) + 0.999999, 0);

						--message 'vo_summa = ', vo_summa to client;
					    for src_vent as sv sensitive cursor for
						select rest as sv_rest, ventureId as sv_srcVentureId 
							from #vntRest vr
							where 
								round(rest - vo_summa, 3) >= 0
								and vr.ventureId != dv_dstVentureId
							order by rest desc
						do
							-- Проверка на "паразитное" добавление после прихода товара на 
							-- фирму у которой не было отрицательного остатка, а у другой фирме
							-- этот отрицательный остаток был. Из-за того, что приход проходит
							-- всегда первым, взаимозачет увеличивается каждый раз при запуске процедуры

							if sec_sort <> 1 then
								--message 'sv_rest = ', sv_rest to client;
								set ivo_id = wf_put_ivo_nomnom (
									  convert(date, r_nDate)
									, p_nomnom
									, vo_summa
									, p_procentOver
									, sv_srcVentureId
									, dv_dstVentureId
									, p_term_start
									, p_term_end
								);
								message '	cоздана/изменена накладная №', ivo_id to client;
								update #vntRest set rest = rest + vo_summa where ventureId = dv_dstVentureId;
								update #vntRest set rest = rest - vo_summa where ventureId = sv_srcVentureId;
								leave compensate;
							end if;
						end for;
					end for;
				end if;
			end if;
/*			
			select rest into rest1 from #vntRest where ventureId = 1;
			select rest into rest2 from #vntRest where ventureId = 2;
			message r_numDoc, '        ', r_nDate,'      ', r_qty, '      ', round(rest1 + rest2, 2), '      ', round(rest1, 2), '      ', round(rest2, 2) 
			, '      ', r_ventureId
			, '      ', r_destventureId
			to client;
*/
		end for;
end;


--------------------------------------------------------------
if exists (select '*' from sysprocedure where proc_name like 'fill_venture_order') then  
	drop procedure fill_venture_order;
end if;

if exists (select '*' from sysprocedure where proc_name like 'ivo_generate') then  
	drop procedure ivo_generate;
end if;

create 
	-- автоматическое формирование взаимозачета между предприятиями
	-- если p_nomnom задан - только для этой номенклатуры. Иначе по всему справочнику номенклатуры.
	-- 
	procedure ivo_generate (
	  p_procentOver float
	, p_term_start date    default null
	, p_term_end date      default null
	, p_nomnom varchar(50) default null
)

begin
	declare v_defaultVentureId integer;
	
	select v.ventureId into v_defaultVentureId 
	from guideventure v 
	join system s on s.id_analytic_default = v.id_analytic;
--	message v_defaultVentureId to client;


	create table #vntRest (ventureId integer, rest float);
		
	insert into #vntRest (ventureId, rest)
	select ventureId, 0.0
	from guideVenture where id_analytic is not null;

	guide_loop:
	for cur_nom as cn sensitive cursor for
		select nomnom as r_nomnom
		from sguidenomenk n 
		where nomnom = isnull(p_nomnom, nomnom)
	do
		call ivo_generate_nomnom (
			  r_nomnom
			, p_procentOver
			, v_defaultVentureId
			, p_term_start
			, p_term_end
		);
--		leave guide_loop;
	end for;
	drop table #vntRest;
end;
		
	
if exists (select '*' from sysprocedure where proc_name like 'ivo_generate_numdoc') then  
	drop procedure ivo_generate_numdoc;
end if;

create 
	-- автоматическое формирование взаимозачета между предприятиями по номенклатуре, входящей в накладную (группу для одного закакза)
	-- 
	procedure ivo_generate_numdoc (
	  in p_numdoc integer
	, p_procentOver float
	, p_term_start date default null
	, p_term_end date default null
)

begin
	declare v_defaultVentureId integer;
	
	select v.ventureId into v_defaultVentureId 
	from guideventure v 
	join system s on s.id_analytic_default = v.id_analytic;
	message 'Взаимозачет дляя накладной №', p_numdoc to client;


	create table #vntRest (ventureId integer, rest float);
		
	insert into #vntRest (ventureId, rest)
	select ventureId, 0.0
	from guideVenture where id_analytic is not null;

	guide_loop:
	for cur_nom as cn sensitive cursor for
		select distinct nomnom as r_nomnom
		from sdocs n
		join sdmc m on m.numdoc = n.numdoc and m.numext = n.numext
		where n.numdoc = p_numdoc 
	do
		call ivo_generate_nomnom (
			  r_nomnom
			, p_procentOver
			, v_defaultVentureId
			, p_term_start
			, p_term_end
		);
--		leave guide_loop;
	end for;
	drop table #vntRest;
end;
		
/***************************************************************
**	КОНЕЦ КОДА ПРОЦЕДУР/ТРИГГЕРОВ, КОТОРЫЕ ПРЕДНАЗНАЧЕНЫ 
**	ДЛЯЯ ФОРМИРОВАНИЯ ВЗАИМОЗАЧЕТОВ
****************************************************************/

	

if exists (select '*' from sysprocedure where proc_name like 'wf_make_invnm') then  
	drop function wf_make_invnm;
end if;


create 
	/* * * * * * * * * * * * * * * * * * * * * * * * * * *
	 * Функция wf_make_invnm используется для получения
	 * такого названия НЕВАРИАНТНОГО ИЗДЕЛИЯ или НОМЕНКЛАТУРЫ,
	 * как оно будет выглядеть в базах Комтех.
	 * В приоре это название не хранится в базе, а составляется
	 * динамически из Cod, NomName, Size при показе в гриде.
	 * В Комтехе это приходится прописывать жестко, как название 
	 * позиции номенклатуры
	 * * * * * * * * * * * * * * * * * * * * * * * * * * */
 	function wf_make_invnm (
	  p_nomname varchar(50) default null
	, p_size varchar(30) default null
	, p_cod varchar(20) default null
) returns varchar(150)
begin
	    if (p_cod is not null and char_length(p_cod) > 0) then
	    	set wf_make_invnm =
	    		+ p_cod + ' ';
	    end if;

	    set wf_make_invnm = wf_make_invnm + p_nomname;
	    if (p_size is not null and char_length(p_size) > 0) then
	    	set wf_make_invnm = wf_make_invnm 
	    		+ ' ' + p_size;
	    end if;
end;

if exists (select '*' from sysprocedure where proc_name like 'wf_make_variant_nm') then  
	drop function wf_make_variant_nm;
end if;

create 
	/* * * * * * * * * * * * * * * * * * * * * * * * * * *
	 * Функция wf_make_variant_nm используется для получения
	 * такого названия ВАРИАНТНОГО ИЗДЕЛИЯ, 
	 * как оно будет выглядеть в базах Комтех.
	 * * * * * * * * * * * * * * * * * * * * * * * * * * */
	function wf_make_variant_nm (
	  p_nomname varchar(50) default null
	, p_size varchar(30) default null
	, p_cod varchar(20) default null
	, p_xprext varchar(20) default null
) returns varchar(150)
begin
	set wf_make_variant_nm = wf_make_invnm(p_nomName, p_size
		, convert(varchar(2), p_xprext) + '/' + p_cod
	);
	
end;


if exists (select '*' from sysprocedure where proc_name like 'wf_retrieve_bill_company') then  
	drop function wf_retrieve_bill_company;
end if;

create function wf_retrieve_bill_company (
	  p_id_bill integer
	, p_ventureName varchar(50)
) returns varchar(150)
begin
	declare v_serverName varchar(20);

	select sysname into v_serverName 
	from GuideVenture where ventureName = p_ventureName;
    
    --message 'sysname = ', v_serverName  to client;

		set wf_retrieve_bill_company = select_remote(
			v_serverName
			, 'voc_names'
			, 'nm'
			, 'id = ' + convert( varchar(20), p_id_bill)
		);
end;



if exists (select '*' from sysprocedure where proc_name like 'wf_check_jscet_split') then  
	drop function wf_check_jscet_split;
end if;

// возвращает id счета (бухгалтерского) из которого будет удаляться
create function wf_check_jscet_split (
	p_numorder integer            // заказ, которому меняем номер счета
) returns integer
begin
	declare remoteServerOld varchar(32);
	declare varchar_id varchar(20);
	declare v_invoice varchar(10);
	declare f_exists integer;

	// аттрибуты заказа который может быть слит с другим
	// (тот у которого руками меняем номер счета)
	declare old_invoice varchar(10);
	declare old_ventureId integer;
	declare old_id_jscet integer;
	declare old_invCode varchar(20);
	declare old_server varchar(20);
	declare old_numorder integer;

	set wf_check_jscet_split = null;

	select numorder, invoice, id_jscet, o.ventureId, v.invCode, v.sysname 
	into old_numorder, old_invoice, old_id_jscet, old_ventureId, old_invCode, old_server
	from orders o
		join guideventure v on v.ventureId = o.ventureId
	where numorder = p_numorder;

	if old_ventureId is null then
		return;
	end if;

	select count(*)
	into wf_check_jscet_split
	from orders o
    where o.invoice = old_invoice
	and o.numorder != old_numorder
	and isnull(o.shipped, 0) = 0
	and substring(o.numorder,0,1) = substring(p_numorder, 0, 1)
	;

end;


-------------------------------------
-------------------------------------
-------------------------------------
if exists (select '*' from sysprocedure where proc_name like 'wf_jscet_handle') then 
	drop function wf_jscet_handle;
end if;

// id бухгалтерского счета для заказа
create function wf_jscet_handle (
	// заказ, который должен быть выделен в отдельный счет
	  p_numorder integer			
	, in p_id_jscet_new integer default null
) returns integer
begin
	// аттрибуты заказа который может быть разделен на два
	declare old_invoice varchar(10);
	declare old_ventureId integer;
	declare old_id_jscet integer;
	declare old_invCode varchar(20);
	declare old_server varchar(20);
	declare v_nu_jscet varchar(50);
	declare v_id_jscet integer;
	declare v_id_dest integer;

	--message 'p_numorder = ', p_numorder to client;
	--message 'p_id_jscet_new = ', p_id_jscet_new to client;

	select invoice, id_jscet, o.ventureId, v.invCode, v.sysname, f.id_voc_names 
	into old_invoice, old_id_jscet, old_ventureId, old_invCode, old_server, v_id_dest
	from orders o
		join guideventure v on v.ventureId = o.ventureId
		join guidefirms f on f.firmid = o.firmid
	where numorder = p_numorder;

	if old_ventureId is null then
		return;
	end if;

	if p_id_jscet_new is not null then
		// делаем перемещение пре
		set v_id_jscet = p_id_jscet_new;
		set v_nu_jscet = select_remote (old_server, 'jscet', 'nu', 'id = ' + convert(varchar(20), p_id_jscet_new));
//		set out_invoice = old_invCode + convert(varchar(20), v_nu_jscet);
	else
		// выделение заказа в отдельный счет
		call put_jscet (v_id_jscet, v_nu_jscet, old_server, p_numOrder, v_id_dest, old_invoice);
	end if;

	update orders set id_jscet = v_id_jscet where numOrder = p_numorder;
	update orders set invoice = old_invCode + convert(varchar(20), v_nu_jscet) where numOrder = p_numorder;

	// Нужно выделить только те детали счета, которые относятся 
	// к заказу и перенести их в новый счет
	--message ' old_server = ', old_server to client;
	--message ' v_id_jscet = ', v_id_jscet to client;
	--message ' p_numOrder = ', p_numOrder to client;
	call wf_move_invoice_detail (old_server, v_id_jscet, p_numOrder);

	// исправить порядковые 
	// номера позиций для нового и старого счета
	call call_remote(old_server, 'slave_renu_scet', v_id_jscet);
	call call_remote(old_server, 'slave_renu_scet', old_id_jscet);

//	return convert(integer, v_nu_jscet);
	return v_id_jscet;
end;



if exists (select '*' from sysprocedure where proc_name like 'wf_split_jscet') then  
	drop function wf_split_jscet;
end if;

// возвращает id бухгалтерского счета для заказа
// 
create function wf_split_jscet (
	// заказ, который должен быть выделен в отдельный счет
	p_numorder integer
	// номер нового счета
	, p_newInvoice varchar(32) default null
) returns varchar(32)
begin
	set wf_split_jscet = wf_jscet_handle(p_numorder);
	if p_newInvoice is not null then
		update orders set invoice = p_newInvoice where numorder = p_numorder;
	end if;
end;


if exists (select '*' from sysprocedure where proc_name like 'wf_move_jscet') then  
	drop function wf_move_jscet;
end if;

// новый номер бухгалтерского счета для заказа
create function wf_move_jscet (
	// номер заказа - источника, который должен быть перемещен 	
	  p_numorder integer
	// id счета, к которому будет присоединен заказ
	, in p_id_jscet_merge integer
) returns varchar(32)
begin
	set wf_move_jscet = wf_jscet_handle(p_numorder, p_id_jscet_merge);
end;


-------------------------------------
-------------------------------------
-------------------------------------

if exists (select '*' from sysprocedure where proc_name like 'put_jscet') then  
	drop procedure put_jscet;
end if;

create procedure put_jscet (
	  out r_id integer
	, out v_nu_jscet varchar(50)
	, in remoteServerNew varchar(20)
	, in p_numOrder integer
	, in p_id_dest integer
	, in p_nu_old varchar(50) default null 
) 
begin
	declare v_fields varchar(255);
	declare v_values varchar(2000);
	declare r_nu varchar(50);
--	declare v_firm_id integer;
	declare v_invCode varchar(10);
--	declare p_id_dest integer;
	declare v_id_schef integer;
	declare v_id_bux integer;
	declare v_id_bank integer;
	declare v_datev varchar(20);
	declare v_id_cur integer;
	declare v_currency_rate float;
	declare v_order_date varchar(20);
	declare v_check_count integer; 
	declare v_id_jscet integer;
	declare v_intInvoice integer;


	select invCode into v_invCode
	from guideVenture where sysname = remoteServerNew;

/*
	set r_id = select_remote(
		remoteServerNew
		, 'jscet'
		, 'max(id)'
	);
*/

	set v_nu_jscet = nextnu_remote(remoteServerNew, 'jscet', p_nu_old);

	set r_id = r_id + 1;
	set v_order_date = convert(varchar(20), now());
	set v_id_cur = system_currency();
	execute immediate 'call slave_currency_rate_' + remoteServerNew + '(v_datev, v_currency_rate, v_order_date, v_id_cur )';
	
	set v_fields =
		 'nu'
--		+ ', id'
		+ ', rem'
		+ ', id_s'
		+ ', dat' 
		+ ', datv' 
		+ ', state'
		+ ', real_days'
		+ ', id_curr'
		+ ', curr'
//		+ ', id_kad1'
//		+ ', id_kad_bux'
//		+ ', id_s_bank'
		;

	--message 'v_fields  = ', v_fields to client;
	
	set v_values = 
		convert(varchar(20), v_nu_jscet)
--		+ ', ' + convert(varchar(20), r_id)
		+ ', ' + convert(varchar(20), p_numOrder)
		+ ', -1'
		+ ', ''''' + convert(varchar(20), v_order_date, 112) + ''''''
		+ ', ''''' + v_datev + ''''''
		+ ', 1'
		+ ', 3'
		+ ', ' + convert(varchar(20), v_id_cur)
		+ ', ' + convert(varchar(20), v_currency_rate)
		
	;

	--message 'v_values  = ', v_values to client;

	if p_id_dest is not null then
		set v_fields = v_fields
			+ ', id_d'
			+ ', id_d_cargo'
		;
		set v_values = v_values	
			+ ', ' + convert(varchar(20), p_id_dest)
			+ ', ' + convert(varchar(20), p_id_dest)
		;
	end if;

	set r_id = insert_count_remote(remoteServerNew, 'jscet', v_fields, v_values);
end;

-------------------------------------
-------------------------------------
-------------------------------------
if exists (select '*' from sysprocedure where proc_name like 'get_jscet_nu') then  
	drop function get_jscet_nu;
end if;
/*
create function get_jscet_nu (
	remoteServerNew varchar(20)
) returns integer
begin
	declare r_nu varchar(50);
	declare r_id integer;

	set r_id = select_remote(
		remoteServerNew
		, 'jscet'
		, 'max(id)'
	);

	set r_nu = select_remote(
		remoteServerNew
		, 'jscet'
		, 'nu'
		, 'id = ' + convert( varchar(20), r_id)
	);
	set get_jscet_nu = convert(integer, r_nu) + 1;
end;
*/
-------------------------------------------------------------------------
--------------             System      ----------------------------
-------------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_update_System' and tname = 'System') then 
	drop trigger System.wf_update_System;
end if;

create TRIGGER wf_update_System before update on
System
referencing old as old_name new as new_name
for each row
begin
	declare v_fields varchar(1000);
	declare v_values varchar(2000);
	declare v_where varchar(1000);
	declare v_id_cur_rate integer;
	declare v_id_cur integer;
	declare v_currency_rate float;
	declare updated integer;

	if update(kurs) then
		if abs(old_name.kurs) != abs(new_name.kurs) then
			-- update remote bases
			set v_currency_rate = abs(new_name.kurs);
			set v_id_cur_rate = old_name.id_cur_rate;

			set v_fields = 'curse';
			set v_values = 
				'''''' + convert(varchar(20), v_currency_rate) + ''''''
			;
			set v_where = 'id='
				+ convert(varchar(20), v_id_cur_rate) 
				+ ' and dat = ''''' + convert(varchar(20), now(), 112) + '''''';
			;

			set updated = update_count_host(
					'cur_rate'
					, v_fields
					, v_values
					, v_where
			);

			if updated = 0 then
				set v_id_cur_rate = get_nextid('cur_rate');
				set v_fields = 'id, id_cur, dat, curse, rem';
				set v_id_cur = old_name.id_cur;
				set v_values = 
					convert(varchar(20), v_id_cur_rate)
					+', ' + convert(varchar(20), v_id_cur)
					+', ''''' + convert(varchar(20), now(), 112) +''''''
					+', ''''' + convert(varchar(20), v_currency_rate) + ''''''
					+', ''''Установлено в Prior'''''
				;
	
				call insert_host('cur_rate', v_fields, v_values);
				set new_name.id_cur_rate = v_id_cur_rate;

			end if;

		end if;
	end if;

end;



-------------------------------------------------------------------------
--------------             common procs      ----------------------------
-------------------------------------------------------------------------


if exists (select '*' from sysprocedure where proc_name like 'extract_invoice_number') then  
	drop function extract_invoice_number;
end if;

create function extract_invoice_number (
	p_invoice varchar(10)         // номер счета заказа
	,p_invCode varchar(10)        // префикс номера счета для предприятия
) returns varchar(10)
begin
	declare v_invoice varchar(10);
	set v_invoice = substring(p_invoice, 1, char_length(p_invCode));

	if p_invCode is null or char_length(p_invCode) = 0 then
		set extract_invoice_number = p_invoice;
	end if;

//	message 'v_invoice = ', v_invoice to client;

	if p_invCode = v_invoice then 
		set extract_invoice_number = substring(p_invoice, char_length(p_invCode)+1);
	end if;
end;



------------------------------------------------------------------------------------
if exists (select '*' from sysprocedure where proc_name like 'wf_check_jscet_merge') then  
	drop function wf_check_jscet_merge;
end if;

create function wf_check_jscet_merge (
	p_numorder integer            // заказ, которому меняем номер счета
	,p_invoice varchar(10)         // новый номер счета заказа
//	,p_oldInvoice varchar(10)      // прежний номер счета заказа м.б. 'счет ?'
) returns integer
begin
	declare remoteServerOld varchar(32);
	declare varchar_id varchar(20);
	declare v_invoice varchar(10);
	declare f_exists integer;

	// аттрибуты заказа который может быть слит с други
	// (тот у которого руками меняем номер счета)
	declare old_invoice varchar(10);      
	declare old_ventureId integer;
	declare old_id_jscet integer;
	declare old_invCode varchar(20);
	declare old_server varchar(20);
	declare old_firmId integer;


	set wf_check_jscet_merge = 0;

	select invoice, id_jscet, o.ventureId, v.invCode, v.sysname, o.firmId
	into old_invoice, old_id_jscet, old_ventureId, old_invCode, old_server, old_firmId
	from orders o
		join guideventure v on v.ventureId = o.ventureId
	where numorder = p_numorder;

	-- Если есть заказ
	select 0 - count(*) into wf_check_jscet_merge 
		from orders o
		where o.invoice = p_invoice
			and o.numorder != p_numorder
			and isnull(o.shipped, 0) = 0
			and o.ventureId = old_ventureId
			and o.id_jscet is not null and o.id_jscet > 0
			and o.firmId <> old_firmId
			-- только для этого года
            and substring(o.numorder,0,1) = substring(p_numorder, 0, 1)
		;

	if wf_check_jscet_merge < 0 then
		return;
	end if;

	if old_ventureId is null then
		return;
	end if;

	a:
	for v_server_name as a dynamic scroll cursor for
		select o.numOrder as r_numOrder
			, o.id_jscet as r_id_jscet
		from orders o
		where o.invoice = p_invoice
			and o.numorder != p_numorder
			and isnull(o.shipped, 0) = 0
			and o.ventureId = old_ventureId
			and o.id_jscet is not null and o.id_jscet > 0
			and o.firmId = old_firmId
			-- только для этого года
            and substring(o.numorder,0,1) = substring(p_numorder, 0, 1)
	do

		set wf_check_jscet_merge = r_id_jscet;
		leave a;
/*
		set v_invoice = extract_invoice_number(v_invoice, old_invCode);

		set varchar_id = select_remote(old_server, 'jscet', 'max(id)', 'nu = ''''' + v_invoice + '''''');
		set wf_check_jscet_merge = convert(integer, varchar_id);
		if r_id_jscet != wf_check_jscet_merge then
			// есть такой заказ, у которого id счета другой
			// а номер такой же, на который мы хотим перевести заказ p_numOrder
			// Ситуация для слияния заказа в один
			set f_exists = 1;
		else
			// ни о чем не говорит. Это нормальная ситуация, 
			// к примеру, сливается третий заказ в один счет
		end if;
*/
	end for;

end;


-------------------------------------
-------------------------------------
-------------------------------------
if exists (select '*' from sysprocedure where proc_name like 'wf_merge_jscet') then  
	drop procedure wf_merge_jscet;
end if;

create procedure wf_merge_jscet (
	  p_numorder integer			// заказ, которому меняем номер счета
	, p_id_jscet_new integer        // id счета бухгалтерской базы
	, p_nu_jscet varchar(32)        // номер бух. счета
)
begin
	declare v_updated integer;
	// аттрибуты заказа который может быть слит с други
	// (тот у которого руками меняем номер счета)
	declare old_invoice varchar(10);      
	declare old_ventureId integer;
	declare old_id_jscet integer;
	declare old_invCode varchar(20);
	declare old_server varchar(20);
	declare scet_nu varchar(20);
	declare v_blank_inv integer;

	
	select invoice, id_jscet, o.ventureId, v.invCode, v.sysname 
	into old_invoice, old_id_jscet, old_ventureId, old_invCode, old_server
	from orders o
		join guideventure v on v.ventureId = o.ventureId
	where numorder = p_numorder;

	if old_ventureId is null then
		return;
	end if;
	set scet_nu = select_remote(
		old_server
		, 'scet'
		, 'max(nu)'
		, 'id_jmat = ' + convert(varchar(20), p_id_jscet_new)
	);
	if scet_nu is not null then
		call update_remote(old_server, 'scet', 'nu'
			, 'nu + ' + convert(varchar(20), scet_nu)
			, 'id_jmat = ' + convert (varchar(20), old_id_jscet)
		);
	end if;

	set v_updated = update_count_remote(old_server, 'scet', 'id_jmat'
		, convert(varchar(20), p_id_jscet_new)
		, 'id_jmat = ' + convert (varchar(20), old_id_jscet)
	);

	--message 'old_id_jscet = ', old_id_jscet to client;

	select count(*) into v_blank_inv from orders where id_jscet = old_id_jscet;

	--message 'v_blank_inv = ', v_blank_inv to client;
	if v_blank_inv = 1 then
		-- Исправление баги: счет не обязательно удалять,
		-- а только если не осталось на него ни одной ссылки
		call delete_remote( old_server, 'jscet', 'id = ' + convert(varchar(20), old_id_jscet));
	end if;

	update orders set id_jscet = p_id_jscet_new where numorder = p_numorder;
	update orders set invoice = p_nu_jscet where numorder = p_numorder;

end;


-------------------------------------
-------------------------------------
-------------------------------------
if exists (select '*' from sysprocedure where proc_name like 'recognize_guide') then  
	drop function recognize_guide;
end if;

create function recognize_guide (
	p_sourId integer
	, p_destId integer
	, p_currency_iso varchar(20) default null
) returns integer
begin
	declare v_id_guide_jmat integer;
	if p_sourId < -1000 and p_destId < -1000 then
		set recognize_guide = 1220;
	elseif p_sourId < -1000 and p_destId >= -1000 then
		-- расход
		if isnull (p_currency_iso, 'RUR') = 'RUR' then
			set recognize_guide = 1210;
		else
			set recognize_guide = 1217;
		end if;
	elseif p_sourId >= -1000 and p_destId < -1000 then
		-- приход
		if isnull (p_currency_iso, 'RUR') = 'RUR' then
			set recognize_guide = 1120;
		else
			set recognize_guide = 1127;
		end if;
	else
		raiserror 17000 'Error in recognize_guide(). Обратитесь к администратору. ';
	end if;

end;


if exists (select '*' from sysprocedure where proc_name like 'gualify_guide') then  
	drop procedure gualify_guide;
end if;

if exists (select '*' from sysprocedure where proc_name like 'qualify_guide') then  
	drop procedure qualify_guide;
end if;

create procedure qualify_guide (
	  p_id_guide_jmat integer
	, out p_tp1 integer
	, out p_tp2 integer
	, out p_tp3 integer
	, out p_tp4 integer
) 
begin
		if p_id_guide_jmat = 1127 then 
		-- приход валютный
			set p_tp1 = 1; set p_tp2 = 1; set p_tp3 = 2; set p_tp4 = 7; 
		elseif p_id_guide_jmat = 1120 then 
		-- приход рублевый
			set p_tp1 = 1; set p_tp2 = 1; set p_tp3 = 2; set p_tp4 = 0;
		elseif p_id_guide_jmat = 1220 then 
		-- межсклад
			set p_tp1 = 2; set p_tp2 = 2; set p_tp3 = 2; set p_tp4 = 0;
		elseif p_id_guide_jmat = 1210 then 
		-- расход
			set p_tp1 = 3; set p_tp2 = 2; set p_tp3 = 1; set p_tp4 = 0; 
		elseif p_id_guide_jmat = 1217 then 
		-- расход в валюте
			set p_tp1 = 3; set p_tp2 = 2; set p_tp3 = 1; set p_tp4 = 7; 
		elseif p_id_guide_jmat = 1023 then 
		-- инвентаризация
			set p_tp1 = 0; set p_tp2 = 0; set p_tp3 = 2; set p_tp4 = 3; 
		end if;
end;


if exists (select '*' from sysprocedure where proc_name like 'get_id_guide_by_key') then  
	drop function get_id_guide_by_key;
end if;

create 
	-- чтобы не запоминать цифры справочников - перевести на мнемонические описания
function get_id_guide_by_key (
	  p_key varchar(20)
	  , p_import integer default null
) returns integer
begin
	if p_key = 'приход' or p_key = 'income' then
		if isnull(p_import, 0) = 0 then
			set get_id_guide_by_key = 1120;
		else
			set get_id_guide_by_key = 1127;
		end if;
	elseif p_key = 'расход' or p_key = 'outcome' then
		if isnull(p_import, 0) = 0 then
			set get_id_guide_by_key = 1210;
		else
			set get_id_guide_by_key = 1217;
		end if;
	elseif p_key = 'инвентаризация' or p_key = 'inventory' then
			set get_id_guide_by_key = 1023;
	elseif p_key = 'межсклад' or p_key = 'intern' then
			set get_id_guide_by_key = 1220;
	end if;
end;


if exists (select '*' from sysprocedure where proc_name like 'wf_get_comtex_tp') then  
	drop function wf_get_comtex_tp;
end if;

create function wf_get_comtex_tp (
	p_id_guide_jmat integer
) returns varchar(20)
begin
	declare v_tp1 integer;
	declare v_tp2 integer;
	declare v_tp3 integer;
	declare v_tp4 integer;

	call qualify_guide (
		  p_id_guide_jmat 	
		, v_tp1 
		, v_tp2 
		, v_tp3 
		, v_tp4 
	);

	set wf_get_comtex_tp = convert(varchar(20), v_tp1)
	                +', '+ convert(varchar(20), v_tp2)
	                +', '+ convert(varchar(20), v_tp3)
	                +', '+ convert(varchar(20), v_tp4)
	;

end;



--------------------------
if exists (select '*' from sysprocedure where proc_name like 'wf_insert_jmat') then  
	drop procedure wf_insert_jmat;
end if;

create procedure wf_insert_jmat (
		p_servername varchar(20)
		, p_id_guide_jmat integer
		, p_id_jmat integer
		, p_jmat_date date
		, p_jmat_nu integer
		, p_osn varchar(100)
		, p_id_currency integer
		, p_datev date
		, p_currency_rate float
		, p_id_s integer
		, p_id_d integer
		, p_id_jscet integer default 0
		, p_id_code integer default 0
)
begin
	declare v_fields varchar(255);
	declare v_values varchar(2000);
	declare v_tp varchar(20);
--	declare out_id integer;
	set v_tp = wf_get_comtex_tp(p_id_guide_jmat);
	set p_id_jscet = isnull(p_id_jscet, 0);
	set p_id_code = isnull(p_id_code, 0);


	set v_fields = 'id'
		+ ', dat'
		+ ' , nu '
		+ ', id_s'
		+ ', id_d'
		+ ', osn'
		+ ', id_guide'
		+ ', tp1, tp2, tp3, tp4'
	;

	if p_id_currency is not null then
		set v_fields = v_fields
			+ ', id_curr'
		;
	end if;
	if p_datev is not null then
		set v_fields = v_fields
			+ ', datv'
		;
	end if;
	if p_currency_rate is not null then
		set v_fields = v_fields
			+ ', curr'
		;
	end if;
	set v_fields = v_fields
		+ ', id_jscet'
		+ ', id_code'

	;   
	set v_values = convert(varchar(20), p_id_jmat)
		+ ', ''''' + convert(varchar(20), p_jmat_date) + ''''''
		+ ', ' + convert(varchar(20), p_jmat_nu)
		+ ', ' + convert(varchar(20), p_id_s)
		+ ', ' + convert(varchar(20), p_id_d)
		+ ', ''''' + p_osn + ''''''
		+ ', ' + convert(varchar(20), p_id_guide_jmat)
		+ ', ' + v_tp
	;
	if p_id_currency is not null then
		set v_values = v_values
			+ ', ' + convert(varchar(20), p_id_currency)
		;
	end if;
	if p_datev is not null then
		set v_values = v_values
			+ ', ''''' + convert(varchar(20), p_datev, 112) + ''''''
		;
	end if;
	if p_currency_rate is not null then
		set v_values = v_values
			+ ', ' + convert(varchar(20), p_currency_rate)
		;
	end if;
	set v_values = v_values
		+ ', ' + convert(varchar(20), p_id_jscet)
		+ ', ' + convert(varchar(20), p_id_code)

	;
	call insert_remote(p_servername, 'jmat', v_fields, v_values);

end;




if exists (select '*' from sysprocedure where proc_name like 'wf_insert_mat') then  
	drop function wf_insert_mat;
end if;

create function wf_insert_mat (
		p_servername varchar(20)
		, p_id_mat integer
		, p_id_jmat integer
		, p_id_inv integer
		, p_mat_nu integer
		, p_quant float
		, p_cena float
		, p_currency_rate float
		, p_id_s integer
		, p_id_d integer
		, p_perList float default 1
--		, p_cenav float
--		, p_date date
--		, p_id_cur integer
--		, p_datev varchar(20)

) returns integer
begin
	declare v_fields varchar(255);
	declare v_values varchar(2000);
	declare v_tp varchar(20);
	declare v_tp1 integer;
	declare v_tp2 integer;
	declare v_tp3 integer;
	declare v_tp4 integer;
	declare v_id_guide integer;

	if p_id_jmat is null then
		raiserror 17002 'wf_insert_mat():: Ошибка в параметрах p_id_jmat == null';
	end if;

	if p_id_mat is null then
		set p_id_mat = get_nextid('mat');
--		execute immediate 'call slave_nextid_' + p_servername + '(''mat'', p_id_mat)';
	end if;


	set wf_insert_mat = p_id_mat;

	set v_id_guide	= select_remote(p_servername, 'jmat', 'id_guide', 'id = ' + convert(varchar(20), p_id_jmat));
	set v_tp1 	= select_remote(p_servername, 'jmat', 'tp1', 'id = ' + convert(varchar(20), p_id_jmat));
	set v_tp2 	= select_remote(p_servername, 'jmat', 'tp2', 'id = ' + convert(varchar(20), p_id_jmat));
	set v_tp3 	= select_remote(p_servername, 'jmat', 'tp3', 'id = ' + convert(varchar(20), p_id_jmat));
	set v_tp4 	= select_remote(p_servername, 'jmat', 'tp4', 'id = ' + convert(varchar(20), p_id_jmat));
	
//	set v_tp = wf_get_comtex_tp(v_id_guide);
	set v_tp = convert(varchar(20), v_tp1)
		+ ',' + convert(varchar(20), v_tp2)
		+ ',' + convert(varchar(20), v_tp3)
		+ ',' + convert(varchar(20), v_tp4)
	;


	set v_fields = 'id'
		+ ', id_jmat'
		+ ', id_inv'
		+ ', nu'
		+ ', id_s'
		+ ', id_d'
		+ ', kol1'
--		+ ', kol3'
--		+ ', kol2'
--		+ ', kol23'
		+ ', tp1, tp2, tp3, tp4'
		+ ', summa'
		+ ', summa_sale'
	;
	if v_id_guide = 1127 then
	--  "приход по импорту в валюте"
		set v_fields = v_fields
			+ ', summav'
			+ ', summa_salev'
		;
	end if;

	set v_values = convert(varchar(20), p_id_mat)
		+ ', ' + convert(varchar(20), p_id_jmat)
		+ ', ' + convert(varchar(20), p_id_inv)
		+ ', ' + convert(varchar(20), p_mat_nu)
		+ ', ' + convert(varchar(20), p_id_s)
		+ ', ' + convert(varchar(20), p_id_d)
		+ ', ' + convert(varchar(20), p_quant / p_perList)
--		+ ', ' + convert(varchar(20), p_quant / p_perList)
--		+ ', ' + convert(varchar(20), p_quant / p_perList)
--		+ ', ' + convert(varchar(20), p_quant / p_perList)
		+ ', ' + v_tp
		+ ', ' + convert(varchar(20), p_quant* p_cena * p_currency_rate / p_perList)
		+ ', ' + convert(varchar(20), p_quant* p_cena * p_currency_rate / p_perList)
	;

	if v_id_guide = 1127 then
	-- приход по импорту в валюте 
		set v_values = v_values 
			+ ', ' + convert(varchar(20), p_quant * p_cena / p_perList)
			+ ', ' + convert(varchar(20), p_quant * p_cena / p_perList)
		;
	end if;
--	message 'Предметы накладной:fields = ', v_fields to client;
--	message '	values = ', v_values to client;
	call insert_remote(p_servername, 'mat', v_fields, v_values);
--	execute immediate 'call slave_insert_'+ p_servername +' (''mat'', ''' +v_fields + ''', ''' + v_values + ''')'

end;


if exists (select '*' from sysprocedure where proc_name like 'wf_insert_scet') then  
	drop procedure wf_insert_scet;
end if;

create function wf_insert_scet (
		  p_servername varchar(20)
		, p_id_jscet integer
		, p_id_inv integer
		, p_quant float
		, p_cena float
		, p_date date
)
returns integer
begin
	declare v_id_scet integer;
	declare v_fields varchar(255);
	declare v_values varchar(2000);
	declare scet_nu integer;
	declare v_currency_rate float;
	declare v_datev varchar(20);
	declare v_id_cur integer;


--	set p_quant = round(p_quant, 2);
--	set p_cena = round(p_cena, 2);

 
  if p_servername is not null and p_id_jscet is not null then
//	execute immediate 'select max(nu)+1 into scet_nu from scet_' + p_servername + ' where id_jmat = ' + convert(varchar(20), p_id_jscet);

	-- Получить следующий порядковый номер счета бух.базы
	set scet_nu = select_remote(
		p_servername
		, 'scet'
		, 'max(nu)+1'
		, 'id_jmat = ' + convert(varchar(20), p_id_jscet)
	);

	set scet_nu = isnull(scet_nu, 1);

	-- По какому курсу, учитывая, что в бухгалтерии только рубли, а в Приоре - УЕ
	set v_id_cur = system_currency();

	execute immediate 'call slave_currency_rate_' + p_servername + '(v_datev, v_currency_rate, p_date, v_id_cur )';
	
	set v_fields = '
		 id_jmat
		,id_inv
		,kol1
		,nu
		,summa_sale
		,summa_salev
	';

	set v_values = 
		convert(varchar(20), p_id_jscet)
		+', '+ convert(varchar(20), p_id_inv)
		+', '+ convert(varchar(20), p_quant)
		+', '+ convert(varchar(20), scet_nu)
		+', '+ convert(varchar(20), round(p_quant*p_cena*v_currency_rate, 2))
		+', '+ convert(varchar(20), round(p_quant*p_cena, 2))
	;
	--message 'p_cena = ', p_cena to client;
	--message 'p_quant = ', p_quant to client;
	--message 'v_values = ', v_values to client;

	-- изменения в бухгалтерской базе данных
	set v_id_scet = insert_count_remote(p_servername, 'scet', v_fields, v_values);

//	execute immediate 'select id into v_id_scet from scet_' + p_servername + ' s where s.id_jmat = p_id_jscet and s.id_inv = p_id_inv';
/*
	set v_id_scet = select_remote(
		p_servername
		, 'scet s'
		, 'id'
		, 's.id_jmat = '+convert(varchar(20), p_id_jscet) + ' and s.id_inv = ' +convert(varchar(20), p_id_inv)
	);
*/
	return v_id_scet;
  end if;
  return null;
	
end;





/************************************************************/
/*                 HOST PROCEDURES                          */
/************************************************************/



if exists (select '*' from sysprocedure where proc_name like 'get_nextid') then
	drop function get_nextid;
end if;

create function get_nextid(table_name varchar(100)) returns integer
/*
	получает следующий свободный id для таблицы table_name с учетом всех
*/
begin
	declare curId integer;
	declare maxId integer;
	set maxId = 0;set curId = 0;
	
  for v_server_name as a dynamic scroll cursor for
	select srvname as cur_server from sys.sysservers s join guideventure v on s.srvname = v.sysname and v.standalone = 0 do
	
	execute immediate 'call slave_nextid_' + cur_server + '('''+table_name+''', curId)';
	if maxId < curId then
		set maxId = curId;
	end if;
  end for;
  -- получение следующего глобального id опирается на таблицу inc_table, где хранятся эти самые id
  call update_host('inc_table', 'next_id', convert(varchar(20), maxId + 1), 'table_nm = ''''' + table_name + '''''');
  return maxId;
end;





/**
 get_server_name() => @server_name 
 процедура должна вызываться один раз из bootstrap_blocking.
*/                 

if exists (select '*' from sysprocedure where proc_name like 'get_server_name') then  
	drop function get_server_name;
end if;

create function get_server_name ()
returns varchar(20) 
begin
	set get_server_name = @@servername;
	if (substring (get_server_name, 1, 3) = 'dev') then
		set get_server_name = 'prior';
	end if;
	 
end;




/************************************************************/
/*                  PRIOR SPECIFIC PROCS                    */
/************************************************************/


if exists (select '*' from sysprocedure where proc_name like 'wf_set_invoice_detail') then  
	drop procedure wf_set_invoice_detail;
end if;


-- Процедура синхронизирует предметы заказа Приора
-- с предметами счета в бухгалтерской базе комтеха
-- Это нужно сделать, если в заказ сначала 
-- добавть предметы, а только потом назначить предприятие,
-- через которую этот заказ должен пройти.
create procedure wf_set_invoice_detail (
			p_servername varchar(20)
			, p_id_jscet integer
			, p_numOrder integer
			, p_date date
)
begin

	declare v_id_scet integer;
	declare v_id_inv integer;
	declare is_variant integer;
	declare v_id_variant integer;
	declare is_uslug integer;
	declare v_quant float;
	declare v_perList float;

	set is_uslug = 1; // предполагаем изначально, что да


	for c_nomenk as n dynamic scroll cursor for
		select 
			  p.nomNom as r_nomNom
			, p.quant as r_quant
			, p.cenaEd as r_cenaEd
		from xPredmetybynomenk p
		where p.numOrder = p_numOrder
	do
	    set is_uslug = 0; -- есть предметы к заказу, значит не услуга

		select id_inv, perList into v_id_inv, v_perList from sGuideNomenk where nomnom = r_nomNom;
		
		set v_id_scet = 
			wf_insert_scet(
				p_servername
				, p_id_jscet
				, v_id_inv
				, r_quant / v_perList
				, r_cenaEd
				, p_date
			);
		update xPredmetyByNomenk set id_scet = v_id_scet where current of n;

	end for;


	for c_izd as i dynamic scroll cursor for
		select 
			  prId as r_prId
			, prExt as r_prExt
			, quant as r_quant
			, cenaEd as r_cenaEd
		from xPredmetyByIzdelia p
		where p.numOrder = p_numOrder
	do

	    set is_uslug = 0; -- есть предметы к заказу, значит не услуга
		select id_inv into v_id_inv from sGuideProducts where prId = r_prId;

		-- смотрим, является ли изделие вариантным?
		
		select count(*) into is_variant from sVariantPower where productId = r_prId;
		if is_variant = 1 then
			-- ищем и/или добавляем вариант в Inv
			set v_id_variant = wf_get_variant_id(p_numOrder, r_prId, r_prExt);
			select id_inv into v_id_inv 
			from sGuideComplect 
			where 
				id_variant = v_id_variant;
		end if;

		set v_id_scet = 
			wf_insert_scet(
				p_servername
				, p_id_jscet
				, v_id_inv
				, r_quant
				, r_cenaEd
				, p_date
			);

		update xPredmetyByIzdelia set id_scet = v_id_scet, id_inv = v_id_inv where current of i;
	end for;  -- цикла по изделиям

	select ordered into v_quant from orders where numorder = p_numOrder;
	if is_uslug = 1 and abs(v_quant) > 0.001 then
		-- ищем товар под названием "услуга"
		select id_inv into v_id_inv from sGuideNomenk where nomNom = 'УСЛ';


		set v_id_scet = 
			wf_insert_scet(
				p_servername
				, p_id_jscet
				, v_id_inv
				, 1 // quant
				, v_quant//r_cenaEd
				, now()//p_date
			);

	end if;


end;




if exists (select '*' from sysprocedure where proc_name like 'wf_move_invoice_detail') then  
	drop procedure wf_move_invoice_detail;
end if;


-- Процедура написана на основе wf_move_invoice_detail (через Copy&Paste)
-- 
-- Только вместо добаления предметов перепривязываем позицию к другому счету
create procedure wf_move_invoice_detail (
	p_servername varchar(20)
	, p_id_jscet_new integer
	, p_numOrder integer
)
begin

	declare is_uslug integer;
	declare v_updated integer;
	declare v_quant float;
	declare v_id_scet integer;
	declare v_id_jscet integer;
	declare v_id_inv integer;

	set is_uslug = 1; // предполагаем изначально, что да

	for c_nomenk as n dynamic scroll cursor for
		select 
			id_scet as r_id_scet
		from xPredmetybynomenk p
		where p.numOrder = p_numOrder
	do
	    set is_uslug = 0; -- есть предметы к заказу, значит не услуга

		set v_updated = update_count_remote(p_servername, 'scet', 'id_jmat'
			, convert(varchar(20), p_id_jscet_new)
			, 'id = ' + convert (varchar(20), r_id_scet)
		);


	end for;


	for c_izd as i dynamic scroll cursor for
		select 
			id_scet as r_id_scet
		from xPredmetyByIzdelia p
		where p.numOrder = p_numOrder
	do

	    set is_uslug = 0; -- есть предметы к заказу, значит не услуга

		set v_updated = update_count_remote(p_servername, 'scet', 'id_jmat'
			, convert(varchar(20), p_id_jscet_new)
			, 'id = ' + convert (varchar(20), r_id_scet)
		);

	end for;  -- цикла по изделиям

	--message 'is_uslug = ', is_uslug to client;
	select ordered into v_quant from orders where numorder = p_numOrder;
	if is_uslug = 1 then
		-- Искать услугу ровно с такой же суммой
		-- относящуюся к старому счету и перепривязываем ее к новому

		select id_jscet into v_id_jscet from orders where numorder = p_numorder;

		-- ищем товар под названием "услуга"
		select id_inv into v_id_inv from sGuideNomenk where nomNom = 'УСЛ';

		--message 'v_id_jscet     = ', v_id_jscet    to client;
		--message 'p_id_jscet_new = ', p_id_jscet_new to client;
		--message 'v_quant        = ', v_quant       to client;
		--message 'v_id_inv       = ', v_id_inv      to client;

		call call_remote(p_servername, 'slave_move_uslug', 
			         convert(varchar(20), v_id_jscet    )
			+ ', ' + convert(varchar(20), p_id_jscet_new)
			+ ', ' + convert(varchar(20), isnull(v_quant, 0)       )
			+ ', ' + convert(varchar(20), v_id_inv      )
		);

/*
		set v_id_scet = select_remote(
			p_servername
			, 'scet'
			, 'id'
			, 'id_jmat = '+ convert(varchar(20), p_id_jscet_new)
				+ ' and id_inv = ' + convert(varchar(20), v_id_inv)
				+ ' and summa_salev = ' + convert(varchar(20), v_quant)
		);

		if v_id_scet is not null then
			set v_updated = update_count_remote(p_servername, 'scet', 'id_jmat'
				, convert(varchar(20), p_id_jscet_new)
				, 'id = ' + convert (varchar(20), v_id_scet)
			);
		end if;


		set v_id_scet = 
			wf_insert_scet(
				p_servername
				, p_id_jscet
				, v_id_inv
				, 1 // quant
				, v_quant//r_cenaEd
				, now()//p_date
			);
*/
	end if;


end;





-- Получить ид единицы измерения. ид является общим на все базы
-- Если такой единицы еще нет, то она добавляется во все базы
if exists (select '*' from sysprocedure where proc_name like 'wf_getEdizmId') then  
	drop procedure wf_getEdizmId;
end if;

create FUNCTION wf_getEdizmId (edizm varchar(100), p_rem varchar(100) default 'created by stime') returns integer
begin
	declare edizmId integer;
	declare v_values varchar(200);
	select id_edizm into edizmId from edizm where name = edizm;
	if edizmId is not null then
		return edizmId;
	end if;

	set edizmId = get_nextId('edizm');
	set v_values = convert(varchar(20), edizmId) 
		+ ', ''''' + edizm + ''''''
		+ ', '''''+p_rem+'''''';

	call insert_host('edizm', 'id, nm,rem', v_values );
	insert into edizm (id_edizm, name) 
	values (edizmId, edizm);
	
	return edizmId;
end;


-- Получить ид размера. ид является общим на все базы
-- Если такога размера еще нет, то создается новый размер
-- и добавляется во все базы
if exists (select '*' from sysprocedure where proc_name like 'wf_getSizeId') then  
	drop procedure wf_getSizeId;
end if;

create FUNCTION wf_getSizeId (sz varchar(100), p_rem varchar(100) default 'created by stime') returns integer
begin
	declare sizeId integer;
	declare v_values varchar(200);

	select id_size into sizeId from size where name = sz;
	if sizeId is not null then
		return sizeId;
	end if;

	set sizeId = get_nextId('size');
	set v_values = convert(varchar(20), sizeId)
		+ ', ''''' + sz + ''''''
		+ ', '''''+p_rem+'''''';

	call insert_host('size', 'id,nm,rem', v_values );
	insert into size (id_size, name)
	values (sizeId, sz);
	return sizeId;
end;




-------------------------------------------------------------------------
--------------             xPredmetyByIzdelia      ----------------------
-------------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_insert_izd' and tname = 'xPredmetyByIzdelia') then 
	drop trigger xPredmetyByIzdelia.wf_insert_izd;
end if;

create TRIGGER wf_insert_izd before insert on
xPredmetyByIzdelia
referencing new as new_name
for each row
begin
	declare v_id_scet integer;
	declare v_id_jscet integer;
	declare v_id_inv integer;
--	declare v_ventureid integer;
	declare remoteServerNew varchar(32);
	declare v_invCode varchar(10);
	declare v_fields varchar(255);
	declare v_values varchar(2000);
	declare v_date date;
 
	select id_jscet, inDate, sysname, invCode 
	into v_id_jscet, v_date, remoteServerNew, v_invcode  
		from orders o
		join GuideVenture v on o.ventureId = v.ventureId and v.standalone = 0
		where numOrder = new_name.numOrder;

	select id_inv into v_id_inv 
		from sGuideProducts where prId = new_name.prId;
  
	if remoteServerNew is not null and v_id_jscet is not null then
		set v_id_scet =	
			wf_insert_scet (
				remoteServerNew
				, v_id_jscet
				, v_id_inv
				, new_name.quant
				, new_name.cenaEd
				, v_date
			);
		set new_name.id_scet = v_id_scet;
		set new_name.id_inv = v_id_inv;
	end if;
end;



if exists (select 1 from systriggers where trigname = 'wf_update_izd' and tname = 'xPredmetyByIzdelia') then 
	drop trigger xPredmetyByIzdelia.wf_update_izd;
end if;

create TRIGGER wf_update_izd before update on
xPredmetyByIzdelia
referencing old as old_name new as new_name
for each row
begin
	declare v_id_scet integer;
	declare v_numorder integer;
	declare v_belong_id integer;
	declare remoteServerNew varchar(32);
	declare v_values varchar(100);
	declare v_fields varchar(200);
	declare v_currency_rate float;
	
	set v_id_scet = old_name.id_scet;
--	set v_numorder = old_name.numOrder;

	select sysname
	into remoteServerNew
		from orders o
		join GuideVenture v on o.ventureId = v.ventureId and v.standalone = 0
		where numOrder = old_name.numOrder;


	if remoteServerNew is not null then
		if update(quant) or update(cenaEd) then
			set v_currency_rate = system_currency_rate();
			call update_remote(remoteServerNew, 'scet', 'summa_sale'
				, convert(varchar(20), v_currency_rate * new_name.quant*new_name.cenaEd)
				, 'id = ' + convert(varchar(20), v_id_scet)
			);
			call update_remote(remoteServerNew, 'scet', 'summa_salev'
				, convert(varchar(20), new_name.quant*new_name.cenaEd)
				, 'id = ' + convert(varchar(20), v_id_scet)
			);
//			call update_remote(remoteServerNew, 'scet', 'summa_sale', convert(varchar(20), new_name.quant*new_name.cenaEd), 'id = ' + convert(varchar(20), v_id_scet));
		end if;
		if update(quant) then
			call update_remote(remoteServerNew, 'scet', 'kol1', convert(varchar(20), new_name.quant), 'id = ' + convert(varchar(20), v_id_scet));
//			call update_remote(remoteServerNew, 'scet', 'kol3', convert(varchar(20), new_name.quant), 'id = ' + convert(varchar(20), v_id_scet));
		end if;
	end if;
  
end;


if exists (select 1 from systriggers where trigname = 'wf_delete_izd' and tname = 'xPredmetyByIzdelia') then 
	drop trigger xPredmetyByIzdelia.wf_delete_izd;
end if;

create TRIGGER wf_delete_izd before delete on
xPredmetyByIzdelia
referencing old as old_name
for each row
begin
	declare remoteServerNew varchar(32);
	select sysname
	into remoteServerNew
		from orders o
		join GuideVenture v on o.ventureId = v.ventureId and v.standalone = 0
		where numOrder = old_name.numOrder;

	if remoteServerNew is not null then
		call delete_remote(remoteServerNew, 'scet', 'id = ' + convert(varchar(20), old_name.id_scet));
	end if;
end;



-------------------------------------------------------------------------
--------------             xPredmetyByNomenk      -----------------------
-------------------------------------------------------------------------
--select * from scet_pm order by id_jmat desc
--select * from xpredmetybynomenk order by 1 desc
--select max(nu)+1  from scet_pm where id_jmat = 13281



if exists (select 1 from systriggers where trigname = 'wf_insert_nomenk' and tname = 'xPredmetyByNomenk') then 
	drop trigger xPredmetyByNomenk.wf_insert_nomenk;
end if;

create TRIGGER wf_insert_nomenk before insert on
xPredmetyByNomenk
referencing new as new_name
for each row
begin
	declare v_id_scet integer;
	declare v_id_jscet integer;
	declare v_id_inv integer;
	declare v_ventureid integer;
	declare remoteServerNew varchar(32);
	declare v_invCode varchar(10);
	declare v_fields varchar(255);
	declare v_values varchar(2000);
	declare scet_nu integer;
	declare v_date date;
	declare v_perList float;

	select id_jscet, ventureId, inDate  into v_id_jscet, v_ventureId, v_date from orders where numOrder = new_name.numOrder;
	select id_inv, perList into v_id_inv, v_perList from sGuideNomenk where nomNom = new_name.nomNom;
	select sysname, invCode into remoteServerNew, v_invcode from GuideVenture where ventureId = v_ventureId;

	if remoteServerNew is not null and v_id_jscet is not null then
	  -- Заказ, который имеет ссылки в бух.базах интеграции
	  -- т.е. уже назначен той, иди другой фирме
		set new_name.id_scet = 
			wf_insert_scet(
				remoteServerNew
				, v_id_jscet
				, v_id_inv
				, new_name.quant / v_perList
				, new_name.cenaEd
				, v_date
			);
	end if;
	  
end;


if exists (select 1 from systriggers where trigname = 'wf_update_nomenk' and tname = 'xPredmetyByNomenk') then 
	drop trigger xPredmetyByNomenk.wf_update_nomenk;
end if;

create TRIGGER wf_update_nomenk before update on
xPredmetyByNomenk
referencing old as old_name new as new_name
for each row
begin
	declare v_id_scet integer;
	declare v_belong_id integer;
	declare remoteServerNew varchar(32);
	declare v_currency_rate float;

	declare v_values varchar(100);
	declare v_fields varchar(200);
	
	set v_id_scet = old_name.id_scet;

	select sysname
	into remoteServerNew
		from orders o
		join GuideVenture v on o.ventureId = v.ventureId and v.standalone = 0
		where numOrder = old_name.numOrder;


	if remoteServerNew is not null then
		if update(quant) or update(cenaEd) then
			set v_currency_rate = system_currency_rate();
			call update_remote(remoteServerNew, 'scet', 'summa_sale'
				, convert(varchar(20), v_currency_rate * new_name.quant*new_name.cenaEd)
				, 'id = ' + convert(varchar(20), v_id_scet)
			);
			call update_remote(remoteServerNew, 'scet', 'summa_salev'
				, convert(varchar(20), new_name.quant*new_name.cenaEd)
				, 'id = ' + convert(varchar(20), v_id_scet)
			);
        end if;
		if update(quant) then
			call update_remote(remoteServerNew, 'scet', 'kol1', convert(varchar(20), new_name.quant), 'id = ' + convert(varchar(20), v_id_scet));
--			call update_remote(remoteServerNew, 'scet', 'kol3', convert(varchar(20), new_name.quant), 'id = ' + convert(varchar(20), v_id_scet));
		end if;
	end if;
	  
end;
	
	
if exists (select 1 from systriggers where trigname = 'wf_delete_nomenk' and tname = 'xPredmetyByNomenk') then 
	drop trigger xPredmetyByNomenk.wf_delete_nomenk;
end if;
    
create TRIGGER wf_delete_nomenk before delete on
xPredmetyByNomenk
referencing old as old_name
for each row
begin
	declare remoteServerNew varchar(32);
	select sysname
	into remoteServerNew
		from orders o
		join GuideVenture v on o.ventureId = v.ventureId and v.standalone = 0
		where numOrder = old_name.numOrder;

	if remoteServerNew is not null then
		call delete_remote(remoteServerNew, 'scet', 'id = ' + convert(varchar(20), old_name.id_scet));
	end if;
end;



-------------------------------------------------------------------------
--------------             xVariantNomenc      --------------------------
-------------------------------------------------------------------------

if exists (select 1 from systriggers where trigname = 'wf_insert_variant' and tname = 'xVariantNomenc') then 
	drop trigger xVariantNomenc.wf_insert_variant;
end if;

create TRIGGER wf_insert_variant before insert on
xVariantNomenc
referencing new as new_name
for each row
begin
	declare v_id_scet integer;
	declare v_id_jscet integer;
	declare v_id_inv integer;
	declare v_ventureid integer;
	declare remoteServerNew varchar(32);
	declare v_invCode varchar(10);
--	declare v_fields varchar(255);
--	declare v_values varchar(2000);
	declare curNo integer;
	declare v_quant float;
	declare v_cenaEd float;
	declare v_total integer;
	declare v_id_variant integer;
 
    -- Сколько строчек уже вставлено?
	select count(*) into curNo 
	from xVariantNomenc 
	where
		numOrder = new_name.numOrder 
		and prid = new_name.prid 
		and prExt = new_name.prExt;

	-- А сколько нужно"?"
	select numgroup into v_total from sVariantPower where productId = new_name.prid;

	-- поскольку триггер не after, а before, сумма должна быть на единицу меньше
	if curNo + 1 != v_total then
		--еще не все строки вариантного изделия добавлены
		-- ждем, когда будут добавлены все!
		return;
	end if;

	-- Ищем (и добавляем автоматом) реализацию варианта
	set v_id_variant= wf_get_variant_Id(
			 new_name.numOrder
			,new_name.prId
			,new_name.prExt
			,new_name.nomNom
		);

	select id_inv into v_id_inv 
	from sguidecomplect 
	where id_variant = v_id_variant;
	
	select 
		quant
		, cenaEd 
		, id_scet
	into v_quant
		, v_cenaEd 
		, v_id_scet
	from xPredmetyByIzdelia i 
	where
		i.numOrder = new_name.numOrder 
		and i.prId = new_name.prId 
		and i.prExt = new_name.prExt
	;


	select id_jscet, ventureId  into v_id_jscet, v_ventureId from orders where numOrder = new_name.numOrder;
--	select id_inv into v_id_inv from sGuideProducts where prId = new_name.prId;
	select sysname, invCode into remoteServerNew, v_invcode from GuideVenture where ventureId = v_ventureId;
  
	if remoteServerNew is not null and v_id_scet is not null then
	-- Заказ, который имеет ссылки в бух.базах интеграции
	-- т.е. уже назначен той, иди другой фирме
		call update_remote(remoteServerNew, 'scet', 'id_inv', convert(varchar(20), v_id_inv), 'id = ' + convert(varchar(20), v_id_scet));
--		update xPredmetyByIzdelia i set id_scet = v_id_scet where
--			i.numOrder = new_name.numOrder and i.prId = new_name.prId and i.prExt = new_name.prExt;
	end if;
	
end;


-------------------------------------------------------------------------
--------------             sGuideKlass      ----------------------------
-------------------------------------------------------------------------

if exists (select 1 from systriggers where trigname = 'wf_insert_klass' and tname = 'sguideklass') then 
	drop trigger sguideklass.wf_insert_klass;
end if;

create TRIGGER wf_insert_klass before insert on
sguideklass
referencing new as new_name
for each row
begin
	declare v_id_inv integer;
	declare v_values varchar(200);
	declare v_belong_id integer;
	set v_id_inv = get_nextid('inv');
	select id_inv into v_belong_id from sguideklass where klassId = new_name.parentKlassId;

	set v_values = convert(varchar(20), v_id_inv)
		+ ', ' + convert(varchar(20), v_belong_id)
		+ ', ''''' + new_name.klassname + ''''''
		+ ', 1'
	;
	
	call insert_host('inv', 'id, belong_id, nm, is_group', v_values);
	set new_name.id_inv=v_id_inv;
/*
	insert into inv (klassid, parentklassid, NM, is_group)
	select 
		new_name.klassid
		, new_name.parentklassid
		, new_name.klassname
		, 1;
*/
end;



if exists (select 1 from systriggers where trigname = 'wf_update_klass' and tname = 'sguideklass') then 
	drop trigger sguideklass.wf_update_klass;
end if;

create TRIGGER wf_update_klass before update on
sguideklass
referencing old as old_name new as new_name
for each row
begin
	declare v_id_inv integer;
	declare v_belong_id integer;

	declare v_values varchar(100);
	declare v_fields varchar(200);
	
	set v_id_inv = old_name.id_inv;
	
  if update(klassname) then
	call update_host('inv', 'nm', '''''' + new_name.klassName + '''''', 'id = ' + convert(varchar(20), v_id_inv));
--    update inv as pi set
--      nm = new_name.klassName where
--      pi.id = old_name.id_inv
  end if;
  if update(parentklassId) then
	select id_inv into v_belong_id from sguideklass where klassid = new_name.parentklassId;
	call update_host('inv', 'belong_id', convert(varchar(20), v_belong_id), 'id = ' + convert(varchar(20), v_id_inv));
--    update inv as pi set
--      belong_id = p.id_Inv
--	from sguideklass p
--	where
--      pi.id = old_name.id_inv
--	and p.klassId = new_name.parentklassId
  end if;
  
end;


if exists (select 1 from systriggers where trigname = 'wf_delete_klass' and tname = 'sGuideKlass') then 
	drop trigger sGuideKlass.wf_delete_klass;
end if;

create TRIGGER wf_delete_klass before delete on
sGuideKlass
referencing old as old_name
for each row
begin
	if old_name.id_inv is not null then
		call delete_host('inv', 'id = ' + convert(varchar(20), old_name.id_inv));
	end if;
--  delete from inv where id = old_name.id_inv;
end;



-------------------------------------------------------------------------
--------------             sGuideSeries      ----------------------------
-------------------------------------------------------------------------


if exists (select 1 from systriggers where trigname = 'wf_insert_seria' and tname = 'sguideseries') then 
	drop trigger sguideseries.wf_insert_seria;
end if;

create TRIGGER wf_insert_seria before insert on
sguideseries
referencing new as new_name
for each row
begin
	declare v_id_inv integer;
	declare v_values varchar(200);
	declare v_belong_id integer;
	set v_id_inv = get_nextid('inv');
	select id_inv into v_belong_id from sguideseries where seriaId = new_name.parentSeriaId;

	set v_values = convert(varchar(20), v_id_inv)
		+ ', ' + convert(varchar(20), v_belong_id)
		+ ', ''''' + new_name.serianame + ''''''
		+ ', 1'
	;
	
	call insert_host('inv', 'id, belong_id, nm, is_group', v_values);
	set new_name.id_inv=v_id_inv;

/*
	insert into inv (seriaid, parentseriaid, NM, is_group)
	select 
		-new_name.seriaid
		, -new_name.parentseriaid
		, new_name.serianame
		, 1;

  set new_name.id_inv=@@id
*/
end;

if exists (select 1 from systriggers where trigname = 'wf_update_seria' and tname = 'sguideseries') then 
	drop trigger sguideseries.wf_update_seria;
end if;

create TRIGGER wf_update_seria before update on
sguideseries
referencing old as old_name new as new_name
for each row
begin
	declare v_id_inv integer;
	declare v_belong_id integer;

	declare v_values varchar(100);
	declare v_fields varchar(200);
	
	set v_id_inv = old_name.id_inv;
	


  if update(serianame) then
	call update_host('inv', 'nm', '''''' + new_name.seriaName + '''''', 'id = ' + convert(varchar(20), v_id_inv));
/*
    update inv as pi set
      nm = new_name.seriaName where
      pi.id = old_name.id_inv
*/
  end if;
  if update(parentSeriaId) then
	select id_inv into v_belong_id from sguideseries where seriaid = new_name.parentseriaId;
	call update_host('inv', 'belong_id', convert(varchar(20), v_belong_id), 'id = ' + convert(varchar(20), v_id_inv));
/*
    update inv as pi set
      belong_id = p.id_Inv
	from sguideseria p
	where
      pi.id = old_name.id_inv
	and p.seriaId = new_name.parentSeriaId
*/
  end if;
end;




if exists (select 1 from systriggers where trigname = 'wf_delete_seria' and tname = 'sGuideSeries') then 
	drop trigger sGuideSeries.wf_delete_seria;
end if;

create TRIGGER wf_delete_seria before delete on
sGuideSeries
referencing old as old_name
for each row
begin
	if old_name.id_inv is not null then
		call delete_host('inv', 'id = ' + convert(varchar(20), old_name.id_inv));
	end if;
end;



-------------------------
--- Шифры затрат::shiz
-------------------------
if exists (select 1 from systriggers where trigname = 'wf_insert_shiz' and tname = 'shiz') then 
	drop trigger shiz.wf_insert_shiz;
end if;

create TRIGGER wf_insert_shiz before insert on
shiz
referencing new as new_name
for each row
begin
	declare v_id integer;
	declare v_fields varchar(500);
	declare v_values varchar(2000);

	set v_id = get_nextid('shiz');
	set v_fields = 'id, nm';
	set v_values = convert(varchar(20), v_id) + ', ' + '''''' + new_name.nm + '''''';

	call insert_host('shiz', v_fields, v_values);
	set new_name.id = v_id;

end;


if exists (select 1 from systriggers where trigname = 'wf_update_shiz' and tname = 'shiz') then 
	drop trigger shiz.wf_update_shiz;
end if;

create TRIGGER wf_update_shiz before update order 1 on
shiz
referencing old as old_name new as new_name
for each row
begin
	if update(nm) then
		call update_host('shiz', 'nm', '''''' + new_name.nm + '''''', 'id = ' + convert(varchar(20), old_name.id));
	end if;

end;


if exists (select 1 from systriggers where trigname = 'wf_delete_shiz' and tname = 'shiz') then 
	drop trigger shiz.wf_delete_shiz;
end if;

create TRIGGER wf_delete_shiz before delete on
shiz
referencing old as old_name
for each row
begin
	if old_name.id is not null then
		call delete_host('shiz', 'id = ' + convert(varchar(20), old_name.id));
	end if;
end;



if exists (select '*' from sysprocedure where proc_name like 'wf_add_shiz') then  
	drop function wf_add_shiz;
end if;


create 
	function wf_add_shiz (
		p_nm varchar(20)
) returns integer
begin
	declare v_id_exists integer;

	if isnull(p_nm, '') = '' then 
		set wf_add_shiz = -1;
        return;
    end if;

	select id into v_Id_exists from shiz where nm = p_nm;
	if v_id_exists is null then
		insert into shiz (nm, is_main_costs) values (p_nm, null);
		select id into wf_add_shiz from shiz where nm = p_nm;
	else 
		set wf_add_shiz = -1;
	end if;

end;


-------------------------------------------------------------------------
--------------             sGuideNomenk      ----------------------------
-------------------------------------------------------------------------


if exists (select 1 from systriggers where trigname = 'wf_insert_gnomenk' and tname = 'sGuideNomenk') then 
	drop trigger sGuideNomenk.wf_insert_gnomenk;
end if;

create TRIGGER wf_insert_gnomenk before insert on
sGuideNomenk
referencing new as new_name
for each row
begin
	declare v_id_inv integer;
	declare v_fields varchar(500);
	declare v_values varchar(2000);
	declare v_belong_id integer;
    declare v_id_edizm1 integer;
    declare v_id_edizm2 integer;
    declare v_id_size integer;

	set v_id_inv = get_nextid('inv');

	select id_inv into v_belong_id from sguideklass where klassId = new_name.KlassId;

	set v_values = convert(varchar(20), v_id_inv)
		+ ', ' + convert(varchar(20), v_belong_id)
		+ ', ''''' + new_name.nomName + ''''''
		+ ', ''''' + new_name.nomnom + ''''''
	;

	set v_fields = 'id, belong_id, nm, nomen';
	if new_name.ed_izmer is not null and length(new_name.ed_izmer) > 0 then
   	  	set v_id_edizm1 = wf_getEdizmId(new_name.ed_izmer);
   	  	set v_fields = v_fields + ', id_edizm2';
   	  	set v_values = v_values + ', '+convert(varchar(20), v_id_edizm1);
   	end if; 

	if new_name.ed_izmer2 is not null and length(new_name.ed_izmer2) > 0 then
	  	set v_id_edizm2 = wf_getEdizmId(new_name.ed_izmer2);
   	  	set v_fields = v_fields + ', id_edizm1';
   	  	set v_values = v_values + ', '+convert(varchar(20), v_id_edizm2);
   	end if; 

	if new_name.size is not null  and length(new_name.size) > 0 then
	  	set v_id_size = wf_getSizeId(new_name.size);
   	  	set v_fields = v_fields + ', id_size';
   	  	set v_values = v_values + ', '+convert(varchar(20), v_id_size);
   	end if; 

	call insert_host('inv', v_fields, v_values);
  set new_name.id_inv=v_id_inv;

end;



if exists (select '*' from sysprocedure where proc_name like 'wf_price_revert') then  
	drop function wf_price_revert;
end if;


create 
-- возвращает цену из истории 
	function wf_price_revert (
		p_nomnom varchar(20)
		, p_prev_cost float default null
) returns float
begin
	declare sv_manager char(1);
	declare v_change_date datetime;
	declare v_cost float;
	begin
		set sv_manager = @manager;
		set @manager = '.';
    
		select max(change_date) into v_change_date from sPriceHistory where nomnom = p_nomnom;
		if v_change_date is not null then
			select cost into v_cost from sPriceHistory where nomnom = p_nomnom and change_date = v_change_date;
			delete from sPriceHistory where nomnom = p_nomnom and change_date = v_change_date;
			update sguideNomenk set cost = v_cost where nomnom = p_nomnom;
			-- возвращаем текущую предыдущую дату
			select max(change_date) into v_change_date from sPriceHistory where nomnom = p_nomnom;
			select cost into wf_price_revert from sPriceHistory where nomnom = p_nomnom and change_date = v_change_date;
		end if;
		set @manager = sv_manager;
	exception when others then
	end;
end;

if exists (select 1 from systriggers where trigname = 'wf_price_history' and tname = 'sGuideNomenk') then 
	drop trigger sGuideNomenk.wf_price_history;
end if;

create TRIGGER wf_price_history before update order 2 on
sGuideNomenk
referencing old as old_name new as new_name
for each row
when (update (cost))
begin
	declare v_changed_by_id tinyint;
	declare no_history char(1);
	if update(cost) and isnull(old_name.cost, 0) != isnull(new_name.cost, 0)  then
	    begin
			select  managId into v_changed_by_id
			from Guidemanag where manag = @manager;
			set no_history = @manager;
	    exception when others then
	    	set v_changed_by_id = null;
	    	set no_history = '.';
	    end;
	    message 'no_hostory = ', no_history to client;
	    if no_history <> '.' then
		    message 'insert inot ' to client;
			insert into sPriceHistory (nomnom, cost, change_date, changed_by_id)
			values ( old_name.nomnom, old_name.cost, now(), v_changed_by_id);
		end if;
	end if;
end;



if exists (select 1 from systriggers where trigname = 'wf_update_gnomenk' and tname = 'sGuideNomenk') then 
	drop trigger sGuideNomenk.wf_update_gnomenk;
end if;

create TRIGGER wf_update_gnomenk before update order 1 on
sGuideNomenk
referencing old as old_name new as new_name
for each row
begin
	declare v_id_inv integer;
    declare v_belong_id integer;
    declare v_id_edizm integer;
    declare v_id_size integer;
    declare v_nomName varchar(50);
    declare v_size varchar(30);
    declare v_cod varchar(20);
    declare v_nm varchar(100);

	declare v_values varchar(100);
	declare v_fields varchar(200);
	
	set v_id_inv = old_name.id_inv;
	
  if update(nomnom) then
	call update_host('inv', 'nomen', '''''' + new_name.nomnom + '''''', 'id = ' + convert(varchar(20), v_id_inv));
  end if;

  if update(ed_Izmer) then
  	set v_id_edizm = wf_getEdizmId(new_name.ed_izmer);
--	select id_edizm into v_ed_izm from edizm where e.name = new_name.ed_izmer;
	call update_host('inv', 'id_edizm2', convert(varchar(20), v_id_edizm), 'id = ' + convert(varchar(20), v_id_inv));
  end if;

  if update(ed_Izmer2) then
  	set v_id_edizm = wf_getEdizmId(new_name.ed_izmer2);
--	select id_edizm into v_ed_izm from edizm where e.name = new_name.ed_izmer;
	call update_host('inv', 'id_edizm1', convert(varchar(20), v_id_edizm), 'id = ' + convert(varchar(20), v_id_inv));
  end if;
  
  if update(klassId) then
	select id_inv into v_belong_id from sguideklass where klassid = new_name.klassId;
	call update_host('inv', 'belong_id', convert(varchar(20), v_belong_id), 'id = ' + convert(varchar(20), v_id_inv));
  end if;
  
  if update(size) or update (cod) or update(nomName) then
  	if (new_name.nomName != old_name.nomName) then
  		set v_nomName = new_name.nomName;
  	else 
  		set v_nomName = old_name.nomName;
  	end if;

  	if (new_name.cod != old_name.cod) then
  		set v_cod = new_name.cod;
  	else 
  		set v_cod = old_name.cod;
  	end if;

  	if (new_name.size != old_name.size) then
  		set v_size = new_name.size;
	  	set v_id_size = wf_getSizeId(new_name.size);
		call update_host('inv', 'id_size', convert(varchar(20), v_id_size), 'id = ' + convert(varchar(20), v_id_inv));
  	else 
  		set v_size = old_name.size;
  	end if;


	set v_nm = wf_make_invnm (v_nomname, v_size, v_cod);
	call update_host('inv', 'nm', '''''' + v_nm + '''''', 'id = ' + convert(varchar(20), v_id_inv));

  end if;


end;


if exists (select 1 from systriggers where trigname = 'wf_delete_gnomenk' and tname = 'sGuideNomenk') then 
	drop trigger sGuideNomenk.wf_delete_gnomenk;
end if;

create TRIGGER wf_delete_gnomenk before delete on
sGuideNomenk
referencing old as old_name
for each row
begin
	if old_name.id_inv is not null then
		call delete_host('inv', 'id = ' + convert(varchar(20), old_name.id_inv));
	end if;
end;

--------------------------------------------------------------------------
--------------             sGuideProducts      ----------------------------
--------------------------------------------------------------------------


if exists (select 1 from systriggers where trigname = 'wf_insert_gproduct' and tname = 'sguideproducts') then 
	drop trigger sguideproducts.wf_insert_gproduct;
end if;

create TRIGGER wf_insert_gproduct before insert on
sguideproducts
referencing new as new_name
for each row
begin
	declare v_id_inv integer;
	declare v_fields varchar(500);
	declare v_values varchar(2000);
	declare v_belong_id integer;
    declare v_id_edizm1 integer;
    declare v_id_size integer;
    declare v_nm varchar(102);


	set v_id_inv = get_nextid('inv');

	select id_inv into v_belong_id from sguideseries where seriaId = new_name.prSeriaId;
  	set v_id_edizm1 = wf_getEdizmId('шт.');

  set v_fields = 
  	  ' id'
  	+ ',belong_id'
  	+ ',nomen'
    + ',nm'
    + ',prc1'
    + ',is_compl'
    + ', id_edizm1'
	;

	set v_nm = wf_make_invnm (new_name.prDescript, new_name.prSize, new_name.prName);

	set v_values = 
				 convert(varchar(20), v_id_inv)
		+ ', ' + convert(varchar(20), v_belong_id)
		+ ', ''''' + new_name.prName + ''''''
		+ ', ''''' + v_nm + ''''''
		+ ', ' + convert(varchar(20), new_name.cena4)
		+ ', 1'
   	  	+ ', '+convert(varchar(20), v_id_edizm1);
	;


	if new_name.prsize is not null and length(new_name.prsize) > 0 then
	  	set v_id_size = wf_getEdizmId(new_name.prsize);
   	  	set v_fields = v_fields + ', id_size';
   	  	set v_values = v_values + ', '+convert(varchar(20), v_id_size);
   	end if; 

	call insert_host('inv', v_fields, v_values);
  set new_name.id_inv=v_id_inv;
	

end;



if exists (select 1 from systriggers where trigname = 'wf_update_gproducts' and tname = 'sGuideProducts') then 
	drop trigger sGuideProducts.wf_update_gproducts;
end if;

create TRIGGER wf_update_gproducts before update on
sGuideProducts
referencing old as old_name new as new_name
for each row
begin
	declare v_id_inv integer;
    declare v_belong_id integer;
    declare v_id_edizm integer;
    declare v_id_size integer;
    declare v_prDescript varchar(50);
    declare v_prSize varchar(30);
    declare v_prName varchar(20);
    declare v_nm varchar(102);
    declare is_variant integer;

	declare v_values varchar(100);
	declare v_fields varchar(200);
	
	set v_id_inv = old_name.id_inv;


  if update(prSize) or update(prName) or update (prDescript) then

	select 1 into is_variant from svariantpower vp where vp.productid = old_name.prId;

	if (new_name.prDescript != old_name.prDescript) then
		set v_prDescript = new_name.prDescript;
	else 
		set v_prDescript = old_name.prDescript;
	end if;
  
	if (new_name.prName != old_name.prName) then
		set v_prName = new_name.prName;
		call update_host('inv', 'nomen', '''''' + new_name.prName + '''''', 'id = ' + convert(varchar(20), v_id_inv));
		if is_variant is not null then
			call update_host('inv', 'nomen', '''''' + new_name.prName + '''''', 'belong_id = ' + convert(varchar(20), v_id_inv));
		end if;
	else 
		set v_prName = old_name.prName;
	end if;
  
	if (new_name.prSize != old_name.prSize) then
		set v_prSize = new_name.prSize;
		set v_id_size = wf_getSizeId(v_prSize);
		call update_host('inv', 'id_size', convert(varchar(20), v_id_size), 'id = ' + convert(varchar(20), v_id_inv));
		if is_variant is not null then
			call update_host('inv', 'id_size', convert(varchar(20), v_id_size), 'belong_id = ' + convert(varchar(20), v_id_inv));
		end if;
	else 
		set v_prSize = old_name.prSize;
	end if;
  
  
	set v_nm = wf_make_invnm (v_prDescript, v_prSize, v_prName);
	call update_host('inv', 'nm', '''''' + v_nm + '''''', 'id = ' + convert(varchar(20), v_id_inv));
	if is_variant is not null then
		
		for aCursor as a dynamic scroll cursor for
			select 
				  xprext as r_xprext
				, id_inv as r_id_inv_variant
			from sguidecomplect g
			where productid = old_name.prid
		do
			set v_nm = wf_make_variant_nm (
				  v_prDescript
				, v_prSize
				, v_prName
				, r_xprext
			);
			call update_host('inv', 'nm', '''''' + v_nm + '''''', 'id = ' + convert(varchar(20), r_id_inv_variant));
			call update_host('inv', 'nomen', '''''' + v_prName + '''''', 'id = ' + convert(varchar(20), r_id_inv_variant));

		end for;
	end if;

  end if;

/*
  if update(prName) then
	call update_host('inv', 'nomen', '''''' + new_name.prName + '''''', 'id = ' + convert(varchar(20), v_id_inv));
  end if;

  if update(prDescript) then
	call update_host('inv', 'nm', '''''' + new_name.prDescript + '''''', 'id = ' + convert(varchar(20), v_id_inv));
  end if;

  if update(prsize) then
  	set v_id_size = wf_getSizeId(new_name.prsize);
		call update_host('inv', 'id_size', convert(varchar(20), v_id_size), 'id = ' + convert(varchar(20), v_id_inv));
  end if;
*/


  if update(seriaId) then
	select id_inv into v_belong_id from sguideseries where seriaId = new_name.prSeriaId;
	call update_host('inv', 'belong_id', convert(varchar(20), v_belong_id), 'id = ' + convert(varchar(20), v_id_inv));
  end if;
  
end;



if exists (select 1 from systriggers where trigname = 'wf_delete_gproducts' and tname = 'sGuideProducts') then 
	drop trigger sGuideProducts.wf_delete_gproducts;
end if;

create TRIGGER wf_delete_gproducts before delete on
sGuideProducts
referencing old as old_name
for each row
begin
	if old_name.id_inv is not null then
		call delete_host('inv', 'id = ' + convert(varchar(20), old_name.id_inv));
	end if;
end;


----------------------------------------------------------------------
--------------             sProducts      ----------------------------
----------------------------------------------------------------------



if exists (select 1 from systriggers where trigname = 'wf_insert_product' and tname = 'sProducts') then 
	drop trigger sProducts.wf_insert_product;
end if;

create TRIGGER wf_insert_product before insert order 1 on
sProducts
referencing new as new_name
for each row
begin

  declare v_table_name varchar(30);
  declare v_values varchar(100);
  declare v_fields varchar(200);
  
  declare v_id_inv integer; -- id номенклатуры
  declare v_id_belong_inv integer; -- id изделия
  declare v_id_compl integer; -- backref
  
  declare is_variant integer; -- проверка того, что изделие простое
  
  declare v_id_edizm integer;
  declare v_edizm varchar(50);
  
  
  update sGuideVariant as gv set c = c+1 where gv.productid = new_name.productId and gv.xgroup = new_name.xgroup;
  if @@rowcount = 0 then
    insert into sGuideVariant(c,productid,xgroup) values(
      1,new_name.productId,new_name.xgroup)
  end if;
  
  select numgroup into is_variant from svariantpower where productid = new_name.productid;

  --if (is_variant is null) then
	//Грузим комплектацию 
    // простое (не вариантное) (пока!) изделие
	set v_table_name = 'compl';
	set v_id_compl = get_nextId (v_table_name);
	select id_inv, ed_izmer into v_id_Inv, v_edizm from sguidenomenk where nomnom = new_name.nomNom;
	set v_id_edizm = wf_getEdizmId (v_edizm);

	select id_inv into v_id_belong_inv from sguideproducts where prid = new_name.productId;
	
	set v_fields ='id'
		+ ', id_inv'
		+ ', id_inv_belong'
		+ ', id_edizm'
		+ ', kol'
		;
	
	set v_values =
			 convert(varchar(20), v_id_compl )
			+ ', ' + convert(varchar(20), v_id_inv)
			+ ', ' + convert(varchar(20), v_id_belong_inv)
			+ ', ' + convert(varchar(20), v_id_edizm)
			+ ', ' + convert(varchar(20), new_name.quantity)
		;	

	call insert_host (v_table_name, v_fields, v_values);
	set new_name.id_compl = v_id_compl;
  --end if;
/*
  insert into compl (id_inv, id_inv_belong, id_edizm, kol)
	select gn.id_inv, gp.id_inv, wf_getEdizmId (gn.ed_izmer), new_name.quantity
  from sguideproducts gp
  join sguidenomenk gn on gn.nomNom = new_name.nomNom
  where gp.prid = new_name.productId;
*/

end;



if exists (select 1 from systriggers where trigname = 'wf_update_product' and tname = 'sProducts') then 
	drop trigger sProducts.wf_update_product;
end if;

create TRIGGER wf_update_product before update on
sProducts
referencing old as old_name new as new_name
for each row
begin
  declare namedFromAfter integer;
  
  if update(xgroup) then
  	update sGuideVariant as gv set c = c-1 where gv.productid = old_name.productId and gv.xgroup = old_name.xgroup;
	select c into namedFromAfter from sGuideVariant gv where gv.productid = old_name.productId and gv.xgroup = old_name.xgroup;
	if namedFromAfter = 0 then
		delete from sGuideVariant where productid = old_name.productId and xgroup = old_name.xgroup;
	end if;

	update sGuideVariant as gv set c = c+1 where gv.productid = old_name.productId and gv.xgroup = new_name.xgroup;
	if @@rowcount = 0 then
	 		insert into sGuideVariant (c, productid, xgroup) 
			values( 1, old_name.productId, new_name.xgroup);
	end if;
  	
  
  end if;

  if update (quantity) then
	call update_host('compl', 'kol', convert(varchar(20), new_name.quantity), 'id = ' + convert(varchar(20), old_name.id_compl))
/*
 	update compl c set kol = new_name.quantity
  	from sguideproducts gp  
  	join sguidenomenk gn on gn.nomNom = old_name.nomNom
  	where gp.prid = old_name.productid 
  	and c.id_inv = gn.id_inv and c.id_inv_belong = gp.id_inv;
*/
  end if;

end;


if exists (select 1 from systriggers where trigname = 'wf_delete_product' and tname = 'sProducts') then 
	drop trigger sProducts.wf_delete_product;
end if;

create TRIGGER wf_delete_product after delete on
sProducts
referencing old as old_name
for each row
begin
    declare namedFromAfter integer;
  	update sGuideVariant as gv set c = c-1 where gv.productid = old_name.productId and gv.xgroup = old_name.xgroup;
	select c into namedFromAfter from sGuideVariant gv where gv.productid = old_name.productId and gv.xgroup = old_name.xgroup;
	if namedFromAfter <= 0 then
		delete from sGuideVariant where productid = old_name.productId and xgroup = old_name.xgroup;
	end if;
	if old_name.id_compl is not null then
		call delete_host('compl', 'id = ' + convert(varchar(20), old_name.id_compl));
	end if;
end;
----------------------------------------------------------------------
--------------             sGuideVariant      ------------------------
----------------------------------------------------------------------



if exists (select 1 from systriggers where trigname = 'wf_insert_gvariant' and tname = 'sGuideVariant') then 
	drop trigger sGuideVariant.wf_insert_gvariant;
end if;

create TRIGGER wf_insert_gvariant after insert on
sGuideVariant
referencing new as new_name
for each row
begin
	-- вроде бы ничего не нужно делать
	-- в штатном режиме при добавлении номенклатуры к изделию
	-- добавиться может только либо строка с пустой xgroup
	-- либо строка со значением счетчика, равной 1
	-- И в том и другом случае состояние вариантности не меняется.
end;


if exists (select 1 from systriggers where trigname = 'wf_update_gvariant' and tname = 'sGuideVariant') then 
	drop trigger sGuideVariant.wf_update_gvariant;
end if;


create TRIGGER wf_update_gvariant before update on
sGuideVariant
referencing old as old_name new as new_name
for each row
begin
	declare v_power integer;
	declare v_fixgroups integer;
	
	if update(c) then
		if old_name.xgroup != '' then
			select c into v_fixgroups from sGuideVariant where productId = old_name.productid and xgroup = '';
			select numgroup into v_power from svariantpower where productid = old_name.productid;
			if old_name.c = 1 and new_name.c = 2 then
				update svariantpower set numgroup = numgroup + 1 where productid = old_name.productid;
				if @@rowcount = 0 then
					-- изделие становится вариантным
					insert into svariantpower (numgroup, productid, fixgroups)
					values (1, old_name.productid, v_fixgroups);
				end if;
			elseif old_name.c = 2 and new_name.c = 1 then
				update svariantpower set numgroup = numgroup - 1 where productid = old_name.productid;
				select numgroup into v_power from svariantpower where productid = old_name.productid;
				if v_power = 0 then
					-- изделие перестает быть вариантным
					delete from svariantpower where productid = old_name.productid;
				end if;
			end if;
		else
			-- апдейтим количество фиксированных компонент (если конечно изделие вариантное)
			update svariantpower set fixgroups = new_name.c where productid = old_name.productid;
		end if;
		
	end if;
	
end;


if exists (select 1 from systriggers where trigname = 'wf_delete_gvariant' and tname = 'sGuideVariant') then 
	drop trigger sGuideVariant.wf_delete_gvariant;
end if;

create TRIGGER wf_delete_gvariant after delete on
sGuideVariant
referencing old as old_name
for each row
begin
end;

----------------------------------------------------------------------
--------------             sVariantPower      ------------------------
----------------------------------------------------------------------



if exists (select 1 from systriggers where trigname = 'wf_insert_vpower' and tname = 'sVariantPower') then 
	drop trigger sVariantPower.wf_insert_vpower;
end if;

create TRIGGER wf_insert_vpower after insert on
sVariantPower
referencing new as new_name
for each row
begin
	declare v_id_inv integer;
	select id_inv into v_id_inv from sguideproducts where prid = new_name.productid;
	call update_host('inv', 'is_group', '1', 'id = ' + convert(varchar(20), v_id_inv));
	call update_host('inv', 'is_compl', '0', 'id = ' + convert(varchar(20), v_id_inv));
end;


if exists (select 1 from systriggers where trigname = 'wf_delete_vpower' and tname = 'sVariantPower') then 
	drop trigger sVariantPower.wf_delete_vpower;
end if;

create TRIGGER wf_delete_vpower after delete on
sVariantPower
referencing old as old_name
for each row
begin
	declare v_id_inv integer;
	select id_inv into v_id_inv from sguideproducts where prid = old_name.productid;
	call update_host('inv', 'is_group', '0', 'id = ' + convert(varchar(20), v_id_inv));
	call update_host('inv', 'is_compl', '1', 'id = ' + convert(varchar(20), v_id_inv));
end;


-------------------------------------------------------------------------
--------------             Orders      ----------------------------
-------------------------------------------------------------------------


if exists (select 1 from systriggers where trigname = 'wf_insert_orders' and tname = 'Orders') then 
	drop trigger Orders.wf_insert_orders;
end if;

create TRIGGER wf_insert_orders before insert on
Orders
referencing new as new_name
for each row
begin
end;


if exists (select 1 from systriggers where trigname = 'wf_update_orders' and tname = 'Orders') then 
	drop trigger Orders.wf_update_orders;
end if;

create TRIGGER wf_update_orders before update order 1 on
Orders
referencing old as old_name new as new_name
for each row
begin
	declare remoteServerOld varchar(32);
	declare remoteServerNew varchar(32);
	declare v_fields varchar(255);
	declare v_values varchar(2000);
	declare v_nu_jscet integer;
	declare r_nu varchar(50);
	declare r_id integer;
--	declare v_firm_id integer;
	declare v_invCode varchar(10);
	declare v_id_dest integer;
	declare v_id_schef integer;
	declare v_id_bux integer;
	declare v_id_bank integer;
	declare v_datev varchar(20);
	declare v_id_cur integer;
	declare v_currency_rate float;
	declare v_order_date varchar(20);
	declare v_check_count integer; 
	declare v_id_jscet integer;
	declare v_id_scet integer;
	declare v_id_inv integer;
	declare v_numorder integer;
	declare v_updated integer;
--	declare v_total_account_date datetime;
	declare sync char(1);
	declare c_status_close_id integer;
	declare v_ivo_procent float;

	set c_status_close_id = 6;  -- закрыт

	select sysname, invCode into remoteServerOld, v_invcode from GuideVenture where ventureId = old_name.ventureId;

	if update(invoice) and remoteServerOld is not null then
		call update_remote(remoteServerOld, 'jscet', 'nu'
				, convert(varchar(20), extract_invoice_number(new_name.invoice, v_invCode))
				, 'id = ' + convert(varchar(20), old_name.id_jscet)
		);
	end if;


	if update(ventureId) then
		if new_name.ventureId = 0 then
			set new_name.ventureid = null;
		end if;
		if isnull(old_name.ventureId, 0) != isnull(new_name.ventureId, 0) then
			if remoteServerOld is not null then
				call delete_remote(remoteServerOld, 'jscet', 'id = ' + convert(varchar(20), old_name.id_jscet));
				call delete_remote(remoteServerOld, 'scet', 'id_jmat = ' + convert(varchar(20), old_name.id_jscet));
				set new_name.invoice = 'счет ?';
				set new_name.id_bill = null;
			end if;
		
			select sysname, invCode into remoteServerNew, v_invcode from GuideVenture where ventureId = new_name.ventureId;
			if remoteServerNew is not null then
		
				set v_numOrder = old_name.numOrder;
				select id_voc_names into v_id_dest from guidefirms where firmid = old_name.firmId;
--				set v_firm_id = old_name.firmId;
				call put_jscet(r_id, v_nu_jscet, remoteServerNew, v_numOrder, v_id_dest, old_name.invoice);
		
				set new_name.id_jscet = r_id;
				set new_name.invoice = v_invCode + convert(varchar(20), v_nu_jscet);
				call wf_set_invoice_detail(remoteServerNew, r_id, new_name.numOrder, v_order_date);
			end if;

			-- исправление расходных накладных, связанных с заказом
			--select total_account into v_total_account_date from system;

			-- это можно делать только для тех заказов, которые после перехода на режим полного учета по предприятиям
			update sdocs set ventureId = new_name.ventureId 
			where 
				sdocs.numdoc = new_name.numorder
				--and xDate >= v_total_account_date
			;
		end if;

	end if;
	if update (firmId) and (old_name.id_bill is null or old_name.id_bill = 0) then
		
		select sysname into remoteServerOld from GuideVenture where ventureId = old_name.ventureId;
		if remoteServerOld is not null then
			select id_voc_names into v_id_dest from guideFirms where firmId = new_name.firmId;
			call block_remote(remoteServerOld, get_server_name(), 'jscet');
			call update_remote(remoteServerOld, 'jscet', 'id_d', convert(varchar(20), v_id_dest ), 'id = ' + convert(varchar(20), old_name.id_jscet));
			call update_remote(remoteServerOld, 'jscet', 'id_d_cargo', convert(varchar(20), v_id_dest ), 'id = ' + convert(varchar(20), old_name.id_jscet));
			call unblock_remote(remoteServerOld, get_server_name(), 'jscet');
		end if;
	end if;
	if update (ordered) then

		set v_id_jscet = old_name.id_jscet;
	
		if remoteServerOld is not null and v_id_jscet is not null then
//			message 'remoteServerOld = ', remoteServerOld to client;
			-- Заказ, который имеет ссылки в бух.базах интеграции
			-- т.е. уже назначен той, иди другой фирме

			-- отследить заказ без предметов
			-- сначала проверяем, что он действительно без них
			select count(*) into v_check_count from xpredmetybynomenk where numorder = old_name.numorder;
			if v_check_count > 0 then
				-- заказ с предметами -> ничего не делаем
				return;
			end if;
	    
			select count(*) into v_check_count from xpredmetybyizdelia where numorder = old_name.numorder;
			if v_check_count > 0 then
				-- заказ с предметами -> ничего не делаем
				return;
			end if;
	    
			-- ищем товар под названием "услуга"
			select id_inv into v_id_inv from sGuideNomenk where nomNom = 'УСЛ';

			-- сначала исходим из того, что такая услуга уже есть.
			-- это может произойти при изменении стоимости закакза.

			if abs(new_name.ordered) < 0.001 then
				call delete_remote(remoteServerOld, 'scet'
					, 'id_jmat = ' + convert(varchar(20), v_id_jscet) + ' and id_inv = ' + convert(varchar(20), v_id_inv)
				);
				return;
			end if;
			set v_updated = update_count_remote(
				remoteServerOld
				,'scet', 'summa_salev'
				, convert(varchar(20), new_name.ordered)
				, 'id_jmat = ' + convert(varchar(20), v_id_jscet) + ' and id_inv = ' + convert(varchar(20), v_id_inv)
			);
			execute immediate 'call slave_currency_rate_' + remoteServerOld + '(v_datev, v_currency_rate, v_order_date, v_id_cur )';
//			message 'v_currency_rate = ', v_currency_rate to client;
			set v_updated = update_count_remote(
				remoteServerOld
				,'scet', 'summa_sale'
				, convert(varchar(20), new_name.ordered * v_currency_rate)
				, 'id_jmat = ' + convert(varchar(20), v_id_jscet) + ' and id_inv = ' + convert(varchar(20), v_id_inv)
			);


//			message 'v_updated = ', v_updated to client;
			if v_updated > 0 then
				-- именно такой случай
				return;
			end if;

//			message 'v_id_jscet = ',v_id_jscet  to client;
//			message 'v_id_inv = ', v_id_inv to client;

			-- первый раз меням это поле => нужно добавить
			set v_id_scet = 
				wf_insert_scet(
					remoteServerOld
					, v_id_jscet
					, v_id_inv
					, 1
					, new_name.ordered
					, old_name.indate
				);
		end if;


	end if;
	if 	update (statusId)
		and new_name.statusId = c_status_close_id
--		and wf_order_closed_comtex(old_name.numorder, remoteServerOld) = 1 
	then
		select ivo_procent into v_ivo_procent from system;
--		set v_numorder = old_name.numorder;
		-- генерить взаимозачеты
		call ivo_generate_numdoc(old_name.numorder, v_ivo_procent);

	end if;

end;


if exists (select '*' from sysprocedure where proc_name like 'wf_order_closed_comtex') then  
	drop procedure wf_order_closed_comtex;
end if;

create 
	function wf_order_closed_comtex(
		  in p_numorder integer
		, in p_sysname varchar(32) default null
	) returns integer
begin
	declare v_orders_table varchar(32);
	declare v_old_statusId integer;
	declare v_old_id_jscet integer;
	declare v_gad_level varchar(8);

	set wf_order_closed_comtex = 1;

	if p_sysname = 'stime' then
		-- для аналитики - не делаем проверку на закрытие.
		return;
	end if;

	select tp into v_orders_table from all_orders where numorder = p_numorder;

	execute immediate 'select id_jscet into v_old_id_jscet '
		+ 'from '+ v_orders_table + ' where numorder = ' + convert(varchar(20), p_numorder);

	if      v_old_id_jscet is not null
		and p_sysname != 'stime' -- только для ПМ и ММ
	then
		-- проверить закрыт ли заказ в бухгалтерии
		set v_gad_level = select_remote(p_sysname, 'guides_access_data', 'access_level', 
			'guide_id = 1005 and data_id = ' + convert(varchar(20), v_old_id_jscet) + ' and access_level = 1 '
		);
		if v_gad_level is null then
			set wf_order_closed_comtex = 0;
			--raiserror 17001 'Нельзя закрыть заказ, до тех пор, пока он не закрыт в Бухгалтерии';
		end if;
	end if;

end;


if exists (select 1 from systriggers where trigname = 'last_modified' and tname = 'orders') then 
	drop trigger orders.last_modified;
end if;

create TRIGGER last_modified before update order 2 on 
orders
referencing old as old_name new as new_name
for each row
begin
	if not update(rowLock) and not update(numorder) and not update(lastModified) and not update(id_bill) then
		set new_name.lastModified = now();
	end if;
end;




if exists (select 1 from systriggers where trigname = 'wf_delete_orders' and tname = 'Orders') then 
	drop trigger Orders.wf_delete_orders;
end if;

create TRIGGER wf_delete_orders before delete on
Orders
referencing old as old_name
for each row
begin
	declare remoteServer varchar(32);
	select sysname into remoteServer from guideventure where ventureId = old_name.ventureId;
	if remoteServer is not null then
		call delete_remote(remoteServer, 'jscet', 'id = ' + convert(varchar(20), old_name.id_jscet));
	end if;
--  delete from inv where id = old_name.id_inv;
end;


if exists (select '*' from sysprocedure where proc_name like 'wf_next_numdoc') then  
	drop procedure wf_next_numdoc;
end if;

create 
	function wf_next_numdoc() returns integer
begin
	declare sys_numdoc integer;
	declare sys_numdoc_c varchar(10);
	declare sys_year_i integer;
	declare sys_mmdd char(4);
	declare sys_number_c varchar(4);
--	declare sys_number_i integer;

	declare now_year_ln integer;
	declare now_date char(6);
	declare now_year_i integer;
	declare now_year char(2);
	declare now_mmdd char(4);
	declare now_m char(1);
	declare v_new_base integer;


	-- по умолчанию в том же дне
	set v_new_base = 0;

	-- locking to prevent the concurrent modification
	update system set lastDocNum = lastDocNum;
	select lastDocNum into sys_numdoc from system;
	set sys_numdoc_c = convert(varchar(10), sys_numdoc);

	set now_date = convert(char(6), now(), 12); -- 050716 yymmdd
	set now_year = substring(now_date, 1, 2);
	set now_year_i = convert(integer, now_year); --5 или 10 если 2010-й год
	set now_year_ln = char_length(convert(char(2), now_year_i)); --1 или 2

	-- Стандарная маска номера YMMDDnn[n..] 
	 
	set sys_year_i = convert(integer, substring(sys_numdoc_c, 1, now_year_ln));
	if (sys_year_i != now_year_i) then
		-- Переход на новый год
		set v_new_base = 1;
		-- Учесть переход с 31.12.2009 на 01.01.2010
		-- изменяется длина шаблона номера счета
		--if sys_year_i = 9 and now_year = 10 then
			--??? set v_year_now = 0;
		--end if;
	end if;

	
	set sys_mmdd = substring (sys_numdoc, now_year_ln + 1, 4);
	set now_m = convert(char(1), 2+convert(integer, convert(char(1), substring(now_date,3,1))));
	set now_mmdd = now_m + substring(now_date, 4, 3);
	if sys_mmdd != now_mmdd then
		set v_new_base = 1;
	end if;

	if v_new_base = 0 then
		set sys_number_c = substring (sys_numdoc_c, now_year_ln + 5);
		set sys_number_c = convert(varchar(3), convert(integer, sys_number_c) + 1);
		if char_length(sys_number_c) = 1 then
			set sys_number_c = '0' + sys_number_c;
		end if;
		set wf_next_numdoc = convert(char(2),sys_year_i) + sys_mmdd + sys_number_c;
	else 
		set wf_next_numdoc = convert(char(2),now_year_i) + now_mmdd + '01';
	end if;

	update system set lastDocNum = wf_next_numdoc;


end;


if exists (select '*' from sysprocedure where proc_name like 'wf_next_numorder') then  
	drop procedure wf_next_numorder;
end if;

create 
	function wf_next_numorder() returns integer
begin
	declare sys_numorder integer;
	declare sys_numorder_c varchar(10);
	declare sys_year_i integer;
	declare sys_mmdd char(4);
	declare sys_number_c varchar(4);
--	declare sys_number_i integer;

	declare now_year_ln integer;
	declare now_date char(6);
	declare now_year_i integer;
	declare now_year char(2);
	declare now_mmdd char(4);
	declare v_new_base integer;


	-- по умолчанию в том же дне
	set v_new_base = 0;

	-- locking to prevent the concurrent modification
	update system set lastPrivatNum = lastPrivatNum;

	select lastPrivatNum into sys_numorder from system;
	set sys_numorder_c = convert(varchar(10), sys_numorder);

	set now_date = convert(char(6), now(), 12); -- 050716 yymmdd
	set now_year = substring(now_date, 1, 2);
	set now_year_i = convert(integer, now_year); --5 или 10 если 2010-й год
	set now_year_ln = char_length(convert(char(2), now_year_i)); --1 или 2

	-- Стандарная маска номера YMMDDnn[n..] 
	 
	set sys_year_i = convert(integer, substring(sys_numorder_c, 1, now_year_ln));
	if (sys_year_i != now_year_i) then
		-- Переход на новый год
		set v_new_base = 1;
		-- Учесть переход с 31.12.2009 на 01.01.2010
		-- изменяется длина шаблона номера счета
		--if sys_year_i = 9 and now_year = 10 then
			--??? set v_year_now = 0;
		--end if;
	end if;

	
	set sys_mmdd = substring (sys_numorder, now_year_ln + 1, 4);
	set now_mmdd = substring (now_date, 3, 4);
	if sys_mmdd != now_mmdd then
		set v_new_base = 1;
	end if;

	if v_new_base = 0 then
		set sys_number_c = substring (sys_numorder_c, now_year_ln + 5);
		set sys_number_c = convert(varchar(3), convert(integer, sys_number_c) + 1);
		if char_length(sys_number_c) = 1 then
			set sys_number_c = '0' + sys_number_c;
		end if;
		set wf_next_numorder = convert(char(2),sys_year_i) + sys_mmdd + sys_number_c;
	else 
		set wf_next_numorder = convert(char(2),now_year_i) + now_mmdd + '01';
	end if;

	update system set lastPrivatNum = wf_next_numorder;

end;


-----------------------------------------------------
--	Функции, для работы с вариантными изделиями -----
-----------------------------------------------------

if exists (select 1 from sysprocedure where proc_name = 'wf_get_variant_Id') then
	drop procedure wf_get_variant_Id;
end if;


CREATE FUNCTION wf_get_variant_Id(
	p_numOrder varchar(50)
	, p_productid integer
	, p_prext integer
	, p_incompleteNomnom varchar(20) default null
)
returns integer
begin
	declare v_variantId integer;
	declare is_ok integer;

	-- курсор пробегает по всем комплектам вариантного изделия
	-- которые раньше уже были созданы
	declare c_product_variants dynamic scroll cursor for
		select id_variant from sguidecomplect
		where productId = p_productId;
	open c_product_variants;
	set is_ok = null;
	set v_variantId = 0;

	all_variants: loop
		fetch c_product_variants into v_variantId;
		if SQLCODE <>0 then 
			leave all_variants;
		end if;
		
		set is_ok = wf_try_variant(v_variantId, p_numOrder, p_productId, p_prExt, p_incompleteNomnom);
		if is_ok is not null then
			leave all_variants;
		end if;
	end loop;
	close c_product_variants;
	if is_ok is null then
		set v_variantId = wf_put_variant(p_numOrder, p_productId, p_prExt, p_incompleteNomnom);
	end if;
	return v_variantId;
end;


if exists (select 1 from sysprocedure where proc_name = 'wf_put_variant') then
	drop procedure wf_put_variant;
end if;

CREATE FUNCTION wf_put_variant(p_numOrder varchar(50), p_productid integer, p_prext integer, p_incompleteNomnom varchar(20) default null)
returns integer
begin
	declare order_nom char(50);
	declare v_variantId integer;
//	declare g_id integer; // Глобальный идентификатор на все сервера 
	declare v_xprext integer;
	declare v_nomNom varchar(30);
	declare v_nomName varchar(100);
	declare v_id_size integer;
	declare v_id_edizm integer;
	declare v_prc1 double;
	declare v_id_compl integer;
	declare v_id_Inv integer;
	declare v_id_Inv_compl integer;
	declare v_kol integer;
	declare v_belong_Id integer;
	declare v_variant_id integer;
	declare v_nm varchar(102);
	declare v_size varchar(30);

	declare v_table_name varchar(100);
	declare v_fields varchar(1000);
	declare v_values varchar(1000);

	declare c_order_nom dynamic scroll cursor for
		select nomNom 
		from xVariantnomenc vn
		where vn.prId = p_productId and vn.prExt = p_prExt and vn.numOrder = p_numOrder

				union

		select nomNom from sproducts p
		where 
			    p.productId = p_productId
			and exists (select 1 from svariantpower vp where vp.productid = p.productid)
			and not exists (select 1 from sguidevariant gv where p.productid = gv.productid and p.xgroup = gv.xgroup and not (gv.xgroup = '' or gv.c = 1))

				union

	    select p_incompleteNomnom 
	    where p_incompleteNomnom is not null
		order by 1;

	select max(xPrExt)into v_xprExt from sguideComplect where productId = p_productId; 
	set v_xPrExt = isnull(v_xPrExt, 0) + 1 ;

	// Здесь нужно вставить добавление во все slave.inv таблицы новый комплект вариантного изделия
	//  v_id_inv - новый вариант вариантного изделия
	//  v_belong_id - id папки, которая объединяет все варианты вариантного издлия
	// -----------------------
	set v_id_inv = get_nextid('inv');

		select 
			  prName as v_nomNom
			, prDescript as v_nomName
			, prSize as v_size
			, s.id_size
			, e1.id_edizm as v_id_edizm
			, n.cena4 as v_prc1
			, n.id_inv as v_belong_id
		into
			  v_nomNom
			, v_nomName
			, v_size
			, v_id_size
			, v_id_edizm
			, v_prc1
			, v_belong_id
		from sguideproducts n
		join sguideseries p on p.seriaid = n.prseriaid
		join edizm e1 on e1.name = 'шт.'
		left join size s on s.name = n.prsize
		where n.prid = p_productid;

		set v_id_size = isnull(v_id_size, 0);
		
		// теперь это изделие обязано быть группой,
		// под которой уже будут собираться все варианты
		call update_host('inv', 'is_group', '1', 'id = ' + convert(varchar(20), v_id_inv));

		set v_nm = wf_make_variant_nm (
			  v_nomName
			, v_size
			, v_nomNom
			, v_xprext
		);
	
		// Добавляем вариант в подгруппу		
		set v_fields ='id'
		+ ', belong_id'
		+ ', nomen'
		+ ', nm'
		+ ', id_edizm1'
		+ ', id_size'
		+ ', prc1'
		+ ', is_compl'
		;
		set v_values =
			 convert(varchar(20), v_id_inv)
			+ ', ' + convert(varchar(20), v_belong_id)
			+ ', ''''' + v_nomnom + ''''''
			+ ', ''''' + v_nm + ''''''
			+ ', ' + convert(varchar(20), v_id_edizm)
			+ ', ' + convert(varchar(20), v_id_size)
			+ ', ''''' + convert(varchar(20), v_prc1) + ''''''
			+ ', 1'
		;	
    
		call insert_host ('inv', v_fields, v_values);
	
	
	// Заглолвок комплекта
	insert into sguidecomplect (productId, xPrExt, id_inv)
		values (p_productId, v_xPrExt, v_id_inv);
		
	set v_variantId = @@identity;
		
	open c_order_nom;
	find: loop
		fetch c_order_nom into order_nom;
		if SQLCODE != 0 then
			leave find;
		end if;
		
		// А здесь в slave.compl
		//  ...
		//
		set v_id_compl = get_nextid('compl');
		
		select n.id_inv 
			, e.id_edizm
			, p.quantity
		into 
			v_id_inv_compl
			, v_id_edizm
			, v_kol
		from sproducts p 
		join sguidenomenk n on n.nomnom = order_nom and p.nomnom = n.nomnom
		join edizm e on e.name = n.ed_izmer
		where p.productid = p_productid;

		
		// Добавляем комплектацию варианта во все бфзы
		set v_fields ='id'
		+ ', id_inv'
		+ ', id_inv_belong'
		+ ', id_edizm'
		+ ', kol'
		;
		set v_values =
			 convert(varchar(20), v_id_compl)
			+ ', ' + convert(varchar(20), v_id_inv_compl)
			+ ', ' + convert(varchar(20), v_id_inv)
			+ ', ' + convert(varchar(20), v_id_edizm)
			+ ', ''''' + convert(varchar(20), v_kol) + ''''''
		;	
    
		call insert_host ('compl', v_fields, v_values);

				
		insert into svariantcomplect (id_variant, nomnom, id_compl)
		values (v_variantId, order_nom, v_id_compl);
	end loop;
	close c_order_nom;
	
	return v_variantId;
	
end;



if exists (select 1 from sysprocedure where proc_name = 'wf_try_variant') then
	drop function wf_try_variant;
end if;

CREATE FUNCTION wf_try_variant(p_id_variant integer, p_numOrder varchar(50), p_productid integer, p_prext integer, p_incompleteNomnom varchar(20) default null) returns integer
begin
	
	declare variant_nom char(50);
	declare order_nom char(50);
	declare is_variant_end integer;
	declare is_order_end integer;
	declare ret integer;
	
	declare c_order_nom dynamic scroll cursor for
		select nomNom 
		from xVariantnomenc vn
		where vn.prId = p_productid and vn.prExt = p_prExt and vn.numOrder = p_numorder
				union
		select nomNom from sproducts p
		where 
			    p.productId = p_productId
			and exists (select 1 from svariantpower vp where vp.productid = p.productid)
			and not exists (select 1 from sguidevariant gv where p.productid = gv.productid and p.xgroup = gv.xgroup)
--			and exists (select 1 from xVariantNomenc vn where vn.prId = p.productId and vn.prId = p_productid and vn.prExt = p_prExt and vn.numOrder = p_numorder)
				union
	    select p_incompleteNomnom 
	    where p_incompleteNomnom is not null
	    order by 1;

	declare c_variant_nom dynamic scroll cursor for
		select nomnom from svariantcomplect vc
		where vc.id_variant = p_id_variant
		order by 1;

	open c_order_nom;
	open c_variant_nom;
	set ret = null;
	find: loop
		set is_order_end = 0;
		fetch c_order_nom into order_nom;
		if SQLCODE != 0 then
			set is_order_end = 1;
		end if;
		set is_variant_end = 0;
		fetch c_variant_nom into variant_nom;
		if SQLCODE != 0 then
			set is_variant_end = 1;
		end if;
		if is_order_end = 1 and is_variant_end = 1 then
			set ret = 1; -- success!
			leave find;
		end if;
		if variant_nom is null or order_nom is null then
			leave find;
		end if;
		if is_order_end = 1 or is_variant_end = 1 or variant_nom != order_nom then
			leave find;
		end if;
	end loop;
	close c_variant_nom;
	close c_order_nom;
	return ret;
end;

if exists (select 1 from sysprocedure where proc_name = 'get_currency_rate_id') then
	drop function get_currency_rate_id;
end if;

if exists (select 1 from sysprocedure where proc_name = 'system_currency') then
	drop function system_currency;
end if;

create function system_currency(
	)
	returns integer
begin
	select id_cur into system_currency from system;
end;


if exists (select 1 from sysprocedure where proc_name = 'system_currency_rate') then
	drop function system_currency_rate;
end if;

create function system_currency_rate(
	)
	returns float
begin
	select abs(kurs) into system_currency_rate from system;
end;




----------------------------------------------------------------------
--------------         xPredmetyByNomenkOut          -----------------
----------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_xPredmetyByNomenkOut_outcome_di' and tname = 'xPredmetyByNomenkOut') then 
	drop trigger xPredmetyByNomenkOut.wf_xPredmetyByNomenkOut_outcome_di;
end if;
/*
create 
	trigger wf_xPredmetyByNomenkOut_outcome_di before delete order 1 on 
xPredmetyByNomenkOut
referencing old as old_name
for each row
begin
--	declare v_id_mat integer;
	declare v_id_jmat integer;
	declare v_sysname varchar(50);

--	set v_id_mat = old_name.id_mat;
	set v_id_jmat = old_name.id_jmat;

	select v.sysname
	into v_sysname
	from orders o
		left join guideventure v on v.ventureid = o.ventureid and v.standalone = 0
	where numorder = old_name.numorder;

	call wf_otgruz_remove (
		v_id_jmat
		,'stime'
	);

	if v_sysname is not null and v_sysname != 'stime' then
		call wf_otgruz_remove (
			v_id_jmat
			,v_sysname
		);

	end if;

		
		
end;
*/




if exists (select 1 from systriggers where trigname = 'wf_xPredmetyByNomenkOut_outcome_ui' and tname = 'xPredmetyByNomenkOut') then 
	drop trigger xPredmetyByNomenkOut.wf_xPredmetyByNomenkOut_outcome_ui;
end if;
/*
create 
	trigger wf_xPredmetyByNomenkOut_outcome_ui before update order 1 on 
xPredmetyByNomenkOut
referencing new as new_name old as old_name
for each row
begin
	declare v_id_mat integer;
	declare v_id_jmat integer;
	declare v_sysname varchar(50);
	declare v_cena float;

	if update(quant) and old_name.quant != new_name.quant then
		set v_id_mat = old_name.id_mat;
		set v_id_jmat = old_name.id_jmat;

		select cenaEd into v_cena from xPredmetybyNomenk where numOrder = new_name.numOrder and nomnom = new_name.nomNom;

		select v.sysname
		into v_sysname
		from orders o
		left join guideventure v on v.ventureid = o.ventureid and v.standalone = 0
		where numorder = old_name.numorder;
		

		call wf_otgruz_quant(
			v_id_mat
			,v_id_jmat
			,new_name.quant
			,v_cena
			,'stime'
		);

		if v_sysname is not null and v_sysname != 'stime' then
			call wf_otgruz_quant(
				v_id_mat
				,v_id_jmat
				,new_name.quant
				,v_cena
				,v_sysname
			);

		end if;


	end if;
end;
*/

if exists (select 1 from systriggers where trigname = 'wf_xPredmetyByNomenkOut_outcome_bi' and tname = 'xPredmetyByNomenkOut') then 
	drop trigger xPredmetyByNomenkOut.wf_xPredmetyByNomenkOut_outcome_bi;
end if;

/*
create 
	trigger wf_xPredmetyByNomenkOut_outcome_bi before insert order 1 on 
xPredmetyByNomenkOut
referencing new as new_name
for each row
begin
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_jmat_nu varchar(20);
	declare v_mat_nu varchar(20);
	declare v_currency_rate float;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_id_source integer;
	declare v_id_dest integer;
--	declare v_osn varchar(100);
	declare v_id_jscet integer;
--	declare v_venture_id integer;
	declare v_firm_id integer;
	declare v_sysname varchar(50);
	declare v_ventureName varchar(100);
	declare v_cena float;
	declare v_cur_otgruz_date date;


	if get_standalone('stime') = 1 then
		call log_warning('Информация об отгрузке по заказу ' + convert(varchar(20), new_name.numorder) + ' не попадает в аналитическую базу stime.');
		return;
	end if;

	select max(id_jmat) into v_id_jmat 
	from xPredmetyByIzdeliaOut 
	where numOrder = new_name.numorder and outDate = new_name.outDate;

	if v_id_jmat is null then
		select max(id_jmat) into v_id_jmat 
		from xPredmetyByNomenkOut 
		where numOrder = new_name.numorder and outDate = new_name.outDate;
	end if;

	select 
		 o.id_jscet
		, isnull(s.id_voc_names, 0)
		, isnull(f.id_voc_names,0)
		, v.ventureName
		, v.sysname
	into  
		 v_id_jscet
		, v_id_source
		, v_id_dest
		, v_ventureName
		, v_sysname
	from orders o
		left join guideventure v on v.ventureid = o.ventureid and v.standalone = 0
		left join guidefirms f on o.firmid = f.firmid
		left join sguidesource s on sourceid = -1001
	where numorder = new_name.numorder;

	
	set v_id_currency = system_currency();
	call slave_currency_rate_stime(v_datev, v_currency_rate);

--	select id_voc_names into v_id_dest from guidefirms where firmid = v_firm_id;
--	    message 'v_id_dest = ', v_id_dest to client;
	-- со склада 1 
	-- ?? хотя по идее нужно бы отгружать со склада готовой продукции
--	select id_voc_names into v_id_source from sguidesource where sourceid = -1001;

	if v_id_jmat is null then
--	    message '---' to client;
		set v_id_jmat = wf_otgruz_jmat(
			new_name.numorder
			, v_id_jscet
--			, v_venture_id
			, new_name.outDate
			, v_id_source
			, v_id_dest
			, v_id_currency
			, v_datev
			, v_currency_rate
			, v_sysname
		);
--		update orders set id_jmat = v_id_jmat where numorder = new_name.numorder;
	end if;

	set v_id_mat = get_nextid('mat');
	call slave_select_stime(v_mat_nu, 'mat', 'max(nu)', 'id_jmat = ' + convert(varchar(20), v_id_jmat));
	set v_mat_nu = convert(varchar(20), convert(integer, isnull(v_mat_nu, 0)) + 1);
	select cenaEd into v_cena from xPredmetybyNomenk where numOrder = new_name.numOrder and nomnom = new_name.nomNom;

	call wf_otgruz_nom(
		  v_id_mat
		, v_id_jmat
		, new_name.nomnom
		, new_name.quant
		, v_cena
		, v_mat_nu
		, v_id_source
		, v_id_dest
		, v_currency_rate
		, v_sysname
	);
	set new_name.id_mat = v_id_mat;
	set new_name.id_jmat = v_id_jmat;

end;
*/


----------------------------------------------------------------------
--------------       Otgruz helpers PROCEDURIES           ------------
----------------------------------------------------------------------
if exists (select 1 from sysprocedure where proc_name = 'wf_otgruz_remove') then
	drop procedure wf_otgruz_remove;
end if;

/*
CREATE procedure wf_otgruz_remove(
	  p_id_jmat integer
	, p_sysname varchar(50)
) 
begin
	execute immediate 'call slave_delete_'+p_sysname+'(''jmat''
		, ''id = '' + convert(varchar(20), '+convert(varchar(20), p_id_jmat)+'))'
	;
	execute immediate 'call slave_delete_'+p_sysname+'(''mat''
		, ''id_jmat = '' + convert(varchar(20), '+convert(varchar(20), p_id_jmat)+'))'
	;
end;
*/


if exists (select 1 from sysprocedure where proc_name = 'wf_otgruz_quant') then
	drop procedure wf_otgruz_quant;
end if;

/*
CREATE procedure wf_otgruz_quant(
	  p_id_mat integer
	, p_id_jmat integer
	, p_quant  float
	, p_cena float
--	, p_currency_rate float
	, p_sysname varchar(50)
) 
begin
	declare v_currency_rate float;


	execute immediate 'call slave_select_'+p_sysname+'(v_currency_rate, ''jmat'', ''curr'', ''id = '' + convert(varchar(20), ' + convert(varchar(20), p_id_jmat) +'))';
--	select v_currency_rate;
	execute immediate 'call slave_update_'+p_sysname+'(''mat''
		, ''kol1''
		, '''+convert(varchar(20), round(p_quant, 2))+'''
		, ''id = '' + convert(varchar(20), '+convert(varchar(20), p_id_mat)+'))';
	execute immediate 'call slave_update_'+p_sysname+'(''mat''
		, ''kol3''
		, '''+convert(varchar(20), round(p_quant, 2))+'''
		, ''id = '' + convert(varchar(20), '+convert(varchar(20), p_id_mat)+'))';
	execute immediate 'call slave_update_'+p_sysname+'(''mat''
		, ''summa_sale''
		, '''+convert(varchar(20), round(p_quant* p_cena * v_currency_rate, 2))+'''
		, ''id = '' + convert(varchar(20), '+convert(varchar(20), p_id_mat)+'))';
	execute immediate 'call slave_update_'+p_sysname+'(''mat''
		, ''summa_salev''
		, '''+convert(varchar(20), round(p_quant* p_cena, 2))+'''
		, ''id = '' + convert(varchar(20), '+convert(varchar(20), p_id_mat)+'))';
end;
*/

if exists (select 1 from sysprocedure where proc_name = 'wf_otgruz_nom') then
	drop procedure wf_otgruz_nom;
end if;
/*

CREATE procedure wf_otgruz_nom(
	  p_id_mat integer
	, p_id_jmat integer
	, p_nomnom varchar(50)
	, p_quant  float
	, p_cena float
	, p_mat_nu varchar(20)
	, p_id_source integer
	, p_id_dest integer
	, p_currency_rate float
	, p_sysname varchar(50) default null
) 
begin
--	declare v_id_jmat integer;
--	declare v_id_mat integer;
--	declare v_mat_nu varchar(20);
--	declare v_currency_rate float;
--	declare v_datev date;
--	declare v_id_currency integer;
	declare v_id_inv integer;
--	declare v_id_source integer;
--	declare v_id_dest integer;
--	declare v_cost float;
	declare v_perList float;

	declare sync char(1);

	select id_inv, perList into v_id_inv, v_perList from sguidenomenk where nomnom = p_nomnom;


--	call call_host('block_table', 'sync, ''prior'', ''mat''');
	call block_remote('stime', get_server_name(), 'mat');
	
	call wf_insert_mat (
		'stime'
		,p_id_mat
		,p_Id_jmat
		,v_id_inv
		,p_mat_nu
		,p_quant 
		,p_cena
		,p_currency_rate
		,p_id_source
		,p_id_dest
		,v_perList
	);
	
	if p_sysname is not null and p_sysname != 'stime' then
		call wf_insert_mat (
			p_sysname
			,p_id_mat
			,p_Id_jmat
			,v_id_inv
			,p_mat_nu
			,p_quant 
			,p_cena
			,p_currency_rate
			,p_id_source
			,p_id_dest
			,v_perList
		);
	
	end if;


	call block_remote('stime', get_server_name(), 'mat');
--	call call_host('unblock_table', 'sync, ''prior'', ''mat''');
	--	set wf_otgruz_nom = v_id_mat;
end;
*/


	
	
if exists (select 1 from sysprocedure where proc_name = 'wf_otgruz_jmat') then
	drop function wf_otgruz_jmat;
end if;

/*
CREATE FUNCTION wf_otgruz_jmat(
	p_numorder integer
	,p_id_jscet integer
--	,p_venture_id integer
	,p_date date
	,p_id_source integer
	,p_id_dest integer
	,p_id_currency integer
	,p_datev date
	,p_currency_rate float
	,p_sysname varchar(50) default null
) 
	returns integer
begin
	
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_jmat_nu varchar(20);
--	declare v_currency_rate float;
--	declare v_datev date;
	declare v_id_currency integer;
--	declare v_id_source integer;
--	declare v_id_dest integer;
	declare v_osn varchar(100);
--	declare v_sysname varchar(50);
--	declare v_ventureName varchar(200);

		set v_id_jmat = get_nextid('jmat');
--		set v_id_currency = system_currency();
--		call slave_currency_rate_stime(v_datev, v_currency_rate);
		set v_jmat_nu = nextnu_remote('stime', 'jmat');
		--select id_voc_names into v_id_source from sguidesource where sourceid = new_name.sourid;
		--select id_voc_names into v_id_dest from sguidesource where sourceid = new_name.destid;

		set v_osn = 'заказ N ' + convert(varchar(20), p_numorder);
	    
		call wf_insert_jmat (
			'stime'
			,'1210' --v_id_guide_jmat
			,v_id_jmat
			,p_date --v_jmat_date
			,v_jmat_nu
			,v_osn
			,p_id_currency
			,p_datev
			,p_currency_rate
			,p_id_source
			,p_id_dest
		);

		if p_sysname is not null and p_sysname != 'stime' then
			call wf_insert_jmat (
				p_sysname
				,'1210' --v_id_guide_jmat
				,v_id_jmat
				,p_date --v_jmat_date
				,v_jmat_nu
				,v_osn
				,p_id_currency
				,p_datev
				,p_currency_rate
				,p_id_source
				,p_id_dest
				,p_id_jscet
			);
		end if;

		set wf_otgruz_jmat = v_id_jmat;

end;
*/


----------------------------------------------------------------------
--------------         xPredmetyByIzdeliaOut          ----------------
----------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_xPredmetyByIzdeliaOut_outcome_di' and tname = 'xPredmetyByIzdeliaOut') then 
	drop trigger xPredmetyByIzdeliaOut.wf_xPredmetyByIzdeliaOut_outcome_di;
end if;

/*
create 
	trigger wf_xPredmetyByIzdeliaOut_outcome_di before delete order 1 on 
xPredmetyByIzdeliaOut
referencing old as old_name
for each row
begin

--	declare v_id_mat integer;
	declare v_id_jmat integer;
	declare v_sysname varchar(50);

--	set v_id_mat = old_name.id_mat;
	set v_id_jmat = old_name.id_jmat;

	select v.sysname
	into v_sysname
	from orders o
		left join guideventure v on v.ventureid = o.ventureid and v.standalone = 0
	where numorder = old_name.numorder;

	call wf_otgruz_remove (
		v_id_jmat
		,'stime'
	);

	if v_sysname is not null and v_sysname != 'stime' then
		call wf_otgruz_remove (
			v_id_jmat
			,v_sysname
		);

	end if;

		
		
end;

*/



if exists (select 1 from systriggers where trigname = 'wf_xPredmetyByIzdeliaOut_outcome_ui' and tname = 'xPredmetyByIzdeliaOut') then 
	drop trigger xPredmetyByIzdeliaOut.wf_xPredmetyByIzdeliaOut_outcome_ui;
end if;

/*
create 
	trigger wf_xPredmetyByIzdeliaOut_outcome_ui before update order 1 on 
xPredmetyByIzdeliaOut
referencing new as new_name old as old_name
for each row
begin

	declare v_id_mat integer;
	declare v_id_jmat integer;
	declare v_sysname varchar(50);
	declare v_cena float;

	if update(quant) and old_name.quant != new_name.quant then
		set v_id_mat = old_name.id_mat;
		set v_id_jmat = old_name.id_jmat;

		select cenaEd into v_cena 
		from xPredmetybyIzdelia pi
		where numOrder = new_name.numOrder 
			and pi.prId = new_name.prId 
			and pi.prExt = new_name.prExt
		;

		select v.sysname
		into v_sysname
		from orders o
		left join guideventure v on v.ventureid = o.ventureid and v.standalone = 0
		where numorder = old_name.numorder;
		

		call wf_otgruz_quant(
			v_id_mat
			,v_id_jmat
			,new_name.quant
			,v_cena
			,'stime'
		);

		if v_sysname is not null and v_sysname != 'stime' then
			call wf_otgruz_quant(
				v_id_mat
				,v_id_jmat
				,new_name.quant
				,v_cena
				,v_sysname
			);

		end if;


	end if;

end;
*/

if exists (select 1 from systriggers where trigname = 'wf_xPredmetyByIzdeliaOut_outcome_bi' and tname = 'xPredmetyByIzdeliaOut') then 
	drop trigger xPredmetyByIzdeliaOut.wf_xPredmetyByIzdeliaOut_outcome_bi;
end if;
/*
create 
	trigger wf_xPredmetyByIzdeliaOut_outcome_bi before insert order 1 on 
xPredmetyByIzdeliaOut
referencing new as new_name
for each row
begin
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_jmat_nu varchar(20);
--	declare v_mat_nu varchar(20);
	declare v_currency_rate float;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_id_source integer;
	declare v_id_dest integer;
--	declare v_osn varchar(100);
	declare v_id_jscet integer;
--	declare v_venture_id integer;
	declare v_firm_id integer;
	declare v_sysname varchar(50);
	declare v_ventureName varchar(100);
	declare v_cena float;
	declare v_cur_otgruz_date date;



--	set v_id_jmat = old_name.id_jmat;
	select max(id_jmat) into v_id_jmat 
	from xPredmetyByIzdeliaOut 
	where numOrder = new_name.numorder and outDate = new_name.outDate;

	if v_id_jmat is null then
		select max(id_jmat) into v_id_jmat 
		from xPredmetyByNomenkOut 
		where numOrder = new_name.numorder and outDate = new_name.outDate;
	end if;

	select 
		 o.id_jscet
		, isnull(s.id_voc_names, 0)
		, isnull(f.id_voc_names,0)
		, v.ventureName
		, v.sysname
	into  
		 v_id_jscet
		, v_id_source
		, v_id_dest
		, v_ventureName
		, v_sysname
	from orders o
		left join guideventure v on v.ventureid = o.ventureid and v.standalone = 0
		left join guidefirms f on o.firmid = f.firmid
		left join sguidesource s on sourceid = -1001
	where numorder = new_name.numorder;

	
	set v_id_currency = system_currency();
	call slave_currency_rate_stime(v_datev, v_currency_rate);

	if v_id_jmat is null then
		set v_id_jmat = wf_otgruz_jmat(
			new_name.numorder
			, v_id_jscet
--			, v_venture_id
			, new_name.outDate
			, v_id_source
			, v_id_dest
			, v_id_currency
			, v_datev
			, v_currency_rate
			, v_sysname
		);
	end if;

--	message 'v_id_jscet = ', v_id_jscet to client;
--	message 'v_id_jmat = ', v_id_jmat to client;
	set v_id_mat = get_nextid('mat');
--	call slave_select_stime(v_mat_nu, 'mat', 'max(nu)', 'id_jmat = ' + convert(varchar(20), v_id_jmat));
--	set v_mat_nu = convert(varchar(20), convert(integer, isnull(v_mat_nu, 0)) + 1);

	select cenaEd 
	into v_cena
	from xPredmetybyIzdelia pi
	where pi.numOrder = new_name.numOrder 
		and pi.prId = new_name.prId 
		and pi.prExt = new_name.prExt;

	call wf_otgruz_izd(
		  v_id_mat
		, v_id_jmat
		, new_name.numOrder
		, new_name.prId
		, new_name.prExt
		, new_name.quant
		, v_cena
--		, v_mat_nu
		, v_id_source
		, v_id_dest
		, v_currency_rate
		, v_sysname
	);
	
	
	set new_name.id_mat = v_id_mat;
	set new_name.id_jmat = v_id_jmat;

end;
*/

if exists (select 1 from sysprocedure where proc_name = 'wf_otgruz_izd') then
	drop procedure wf_otgruz_izd;
end if;
/*

CREATE procedure wf_otgruz_izd(
	  p_id_mat integer
	, p_id_jmat integer
	, p_numOrder integer
	, p_prId integer
	, p_prExt integer
	, p_quant  float
	, p_cena float
--	, p_mat_nu varchar(20)
	, p_id_source integer
	, p_id_dest integer
	, p_currency_rate float
	, p_sysname varchar(50) default null
) 
begin
--	declare v_id_jmat integer;
--	declare v_id_mat integer;
	declare v_mat_nu varchar(20);
--	declare v_currency_rate float;
--	declare v_datev date;
--	declare v_id_currency integer;
	declare v_id_inv integer;
--	declare v_id_source integer;
--	declare v_id_dest integer;
--	declare v_cost float;
--	declare v_quant float;

	declare sync char(1);

	
--	call call_host('block_table', 'sync, ''prior'', ''mat''');

	call slave_select_stime(v_mat_nu, 'mat', 'max(nu)', 'id_jmat = ' + convert(varchar(20), p_id_jmat));
--	set v_mat_nu = convert(varchar(20), convert(integer, isnull(v_mat_nu, 0)) + 1);
   
	for aCursor as a dynamic scroll cursor for
	
		select id_inv r_id_inv, cost as r_cost, pr.quantity r_quant, nom.perList as r_perList
		from sguidenomenk nom
		join (
			select nomnom, quantity 
			from sproducts p
			where p.productId = p_prId 
			and ( exists (
					select 1 from sguidevariant vp 
					where 
						p.productid = vp.productid and p.xgroup = vp.xgroup and vp.c = 1
					)
					or p.xgroup = '' or p.xgroup is null
				)
					union 
			select v.nomnom, p.quantity 
			from xvariantnomenc v
				join sproducts p on p.productid = v.prid and p.nomnom = v.nomnom
			where v.prid = p_prId and v.numorder = p_numOrder and v.prExt=p_prExt
		) pr on pr.nomnom = nom.nomnom
	do
	
		
		set v_mat_nu = convert(varchar(20), convert(integer, isnull(v_mat_nu, 0)) + 1);
		call wf_insert_mat (
			'stime'
			,p_id_mat
			,p_Id_jmat
			,r_id_inv
			,v_mat_nu
			,p_quant * r_quant
			,r_cost
			,p_currency_rate
			,p_id_source
			,p_id_dest
			,r_perList
		);
		set p_id_mat = p_id_mat + 1;

	end for;

	
	if p_sysname is not null and p_sysname != 'stime' then
	
--	message 'p_prId = ', p_prId to client;
		if not exists (select 1 from svariantpower where productId = p_prid) then
			select id_inv into v_id_inv from sguideproducts where prId = p_prId;
		else 
			select id_inv into v_Id_inv 
			from xPredmetyByIzdelia pi 
			where pi.prId = p_prId 
				and pi.prExt = p_prExt
				and pi.numOrder = p_numOrder
		end if;
--	message 'v_id_inv = ', v_id_inv to client;

		execute immediate 'call slave_select_'+p_sysname+'(v_mat_nu, ''mat'', ''max(nu)'''
			+', ''id_jmat = '' + convert(varchar(20), '+convert(varchar(20), p_id_jmat)+'))';

		set v_mat_nu = convert(varchar(20), convert(integer, isnull(v_mat_nu, 0)) + 1);

		call wf_insert_mat (
			p_sysname
			,p_id_mat
			,p_Id_jmat
			,v_id_inv
			,v_mat_nu
			,p_quant 
			,p_cena
			,p_currency_rate
			,p_id_source
			,p_id_dest
			,1
		);
	
	end if;


--	call call_host('unblock_table', 'sync, ''prior'', ''mat''');
	--	set wf_otgruz_izd = v_id_mat;
end;
*/
	



-------------------------------------------------------------------------
--------------             BayOrders      ----------------------------
-------------------------------------------------------------------------

if exists (select 1 from systriggers where trigname = 'wf_insert_orders' and tname = 'BayOrders') then 
	drop trigger BayOrders.wf_insert_orders;
end if;

create TRIGGER wf_insert_orders before insert on
BayOrders
referencing new as new_name
for each row
begin
end;


if exists (select 1 from systriggers where trigname = 'wf_update_orders' and tname = 'BayOrders') then 
	drop trigger BayOrders.wf_update_orders;
end if;

create TRIGGER wf_update_orders before update on
BayOrders
referencing old as old_name new as new_name
for each row
begin
	declare remoteServerOld varchar(32);
	declare remoteServerNew varchar(32);
	declare v_fields varchar(255);
	declare v_values varchar(2000);
	declare v_nu_jscet integer;
	declare r_nu varchar(50);
	declare r_id integer;
--	declare v_firm_id integer;
	declare v_invCode varchar(10);
	declare v_id_dest integer;
	declare v_id_schef integer;
	declare v_id_bux integer;
	declare v_id_bank integer;
	declare v_datev varchar(20);
	declare v_id_cur integer;
	declare v_currency_rate float;
	declare v_inv_date varchar(20);
	declare v_numOrder integer;
	declare sync char(1);
	declare c_status_close_id integer;
	declare v_ivo_procent float;

	set c_status_close_id = 6;  -- закрыт

	if update(ventureId) then
		if new_name.ventureId = 0 then
			set new_name.ventureid = null;
		end if;
		if isnull(old_name.ventureId, 0) != isnull(new_name.ventureId, 0) then
			select sysname into remoteServerOld from GuideVenture where ventureId = old_name.ventureId;
			if remoteServerOld is not null then
				call delete_remote(remoteServerOld, 'jscet', 'id = ' + convert(varchar(20), old_name.id_jscet));
				call delete_remote(remoteServerOld, 'scet', 'id_jmat = ' + convert(varchar(20), old_name.id_jscet));
				set new_name.invoice = 'счет ?';
			end if;

			select sysname, invCode into remoteServerNew, v_invcode from GuideVenture where ventureId = new_name.ventureId;

			--message 'sysname = ', remoteServerNew to client;
			
			if remoteServerNew is not null then
	
				set v_numOrder = old_name.numOrder;
--				set v_firm_id = old_name.firmId;
				select id_voc_names into v_id_dest from bayguidefirms where firmid = old_name.firmId;
				call put_jscet(r_id, v_nu_jscet, remoteServerNew, v_numOrder, v_id_dest, old_name.invoice);
		
				set new_name.id_jscet = r_id;
				set new_name.invoice = v_invCode + convert(varchar(20), v_nu_jscet);
				call wf_set_bay_detail(remoteServerNew, r_id, new_name.numOrder, v_inv_date);
			end if;
		end if;
	end if;
	if update (firmId) then
		select sysname into remoteServerOld from GuideVenture where ventureId = old_name.ventureId;
		if remoteServerOld is not null then
			select id_voc_names into v_id_dest from BayGuideFirms where firmId = new_name.firmId;
			call block_remote(remoteServerOld, get_server_name(), 'jscet');
			call update_remote(remoteServerOld, 'jscet', 'id_d', convert(varchar(20), v_id_dest ), 'id = ' + convert(varchar(20), old_name.id_jscet));
			call update_remote(remoteServerOld, 'jscet', 'id_d_cargo', convert(varchar(20), v_id_dest ), 'id = ' + convert(varchar(20), old_name.id_jscet));
			call unblock_remote(remoteServerOld, get_server_name(), 'jscet');
		end if;
	end if;
	if 	update (statusId)
		and new_name.statusId = c_status_close_id
--		and wf_order_closed_comtex(old_name.numorder, remoteServerOld) = 1 
	then
		select ivo_procent into v_ivo_procent from system;
--		set v_numorder = old_name.numorder;
		-- генерить взаимозачеты
		call ivo_generate_numdoc(old_name.numorder, v_ivo_procent);

	end if;
end;


if exists (select 1 from systriggers where trigname = 'wf_delete_orders' and tname = 'BayOrders') then 
	drop trigger BayOrders.wf_delete_orders;
end if;

create TRIGGER wf_delete_orders before delete on
BayOrders
referencing old as old_name
for each row
begin
	declare remoteServer varchar(32);
	select sysname into remoteServer from guideventure where ventureId = old_name.ventureId;
	if remoteServer is not null then
		call delete_remote(remoteServer, 'jscet', 'id = ' + convert(varchar(20), old_name.id_jscet));
	end if;
--  delete from inv where id = old_name.id_inv;
end;



-- Процедура синхронизирует предметы bay-заказа Приора
-- с предметами счета в бухгалтерской базе комтеха
-- Это нужно сделать, если в заказ сначала 
-- добавть предметы, а только потом назначить предприятие,
-- через которую этот заказ должен пройти.
if exists (select '*' from sysprocedure where proc_name like 'wf_set_bay_detail') then  
	drop procedure wf_set_bay_detail;
end if;

create procedure wf_set_bay_detail (
			p_servername varchar(20)
			, p_id_jscet integer
			, p_numOrder integer
			, p_date date
)
begin

	declare v_id_scet integer;
	declare v_id_inv integer;
	declare is_variant integer;
	declare v_id_variant integer;
	declare v_quant float;

	for c_nomenk as n dynamic scroll cursor for
		select 
			  p.nomNom as r_nomNom
			, p.quantity as r_quantity
			, intQuant as r_cenaEd
		from sDmcRez p
		where p.numDoc = p_numOrder
	do

		select 
			n.id_inv
			, r_quantity/n.perList
		into 
			v_id_inv
			, v_quant
		from 
			sGuideNomenk n
		where
			n.nomNom = r_nomNom;


		set v_id_scet = 
			wf_insert_scet(
				p_servername
				, p_id_jscet
				, v_id_inv
				, v_quant
				, r_cenaEd
				, p_date
			);
		update sDmcRez set id_scet = v_id_scet where current of n;

	end for;



end;



-------------------------------------------------------------------------
-------------------             sDmcRez          ------------------------
-------------------------------------------------------------------------
--select * from scet_pm order by id_jmat desc
--select * from sDmcRez order by 1 desc
--select max(nu)+1  from scet_pm where id_jmat = 13281



if exists (select 1 from systriggers where trigname = 'wf_insert_nomenk' and tname = 'sDmcRez') then 
	drop trigger sDmcRez.wf_insert_nomenk;
end if;

create TRIGGER wf_insert_nomenk before insert on
sDmcRez
referencing new as new_name
for each row
begin
	declare v_id_scet integer;
	declare v_id_jscet integer;
	declare v_id_inv integer;
	declare remoteServerNew varchar(32);
	declare v_invCode varchar(10);
	declare v_date date;
	declare v_cenaEd float;
	declare v_quantity float;
	declare v_perList float;


--	message 'sDmcRez.wf_insert_nomenk' to client;
	select 
		o.id_jscet, o.inDate  
		, v.sysname, v.invCode
		, n.id_inv, n.perList 
	into 
		v_id_jscet, v_date 
		, remoteServerNew, v_invcode
		, v_id_inv, v_perList 
	from BayOrders o
	left join GuideVenture v on v.ventureid = o.ventureid and v.standalone = 0
	join sGuideNomenk n on n.nomNom = new_name.nomNom
	where o.numOrder = new_name.numDoc;


	set v_cenaEd = new_name.intQuant;
	set v_quantity = new_name.quantity/v_perList;

--	select id_inv into v_id_inv from sGuideNomenk where nomNom = new_name.nomNom;

--	select sysname, invCode into remoteServerNew, v_invcode from GuideVenture where ventureId = v_ventureId;

	if remoteServerNew is not null and v_id_jscet is not null then
	  -- Заказ, который имеет ссылки в бух.базах интеграции
	  -- т.е. уже назначен той, иди другой фирме
		set new_name.id_scet = 
			wf_insert_scet(
				remoteServerNew
				, v_id_jscet
				, v_id_inv
				, v_quantity
				, v_cenaEd
				, v_date
			);
	end if;
	  
end;


if exists (select 1 from systriggers where trigname = 'wf_update_nomenk' and tname = 'sDmcRez') then 
	drop trigger sDmcRez.wf_update_nomenk;
end if;

create TRIGGER wf_update_nomenk before update on
sDmcRez
referencing old as old_name new as new_name
for each row
begin
	declare v_id_scet integer;
	declare remoteServerNew varchar(32);

	declare v_cenaEd float;
	declare v_quantity float;
	declare v_perList float;
	declare v_currency_rate float;
	
	set v_id_scet = old_name.id_scet;
	  
	select v.sysname
		, n.perList 
	into remoteServerNew
		, v_perList 
	from BayOrders o
	join GuideVenture v on o.ventureId = v.ventureId and v.standalone = 0
	join sGuideNomenk n on n.nomNom = old_name.nomNom
	where numOrder = old_name.numDoc;


	if remoteServerNew is not null then
		if update(quantity) or update(intQuant) then
			set v_currency_rate = system_currency_rate();
			set v_quantity = round(new_name.quantity/v_perList, 2);
			call update_remote(remoteServerNew, 'scet', 'summa_sale'
				, convert(varchar(20), v_currency_rate * v_quantity * new_name.intQuant)
				, 'id = ' + convert(varchar(20), v_id_scet)
			);
			call update_remote(remoteServerNew, 'scet', 'summa_salev'
				, convert(varchar(20), v_quantity*new_name.intQuant)
				, 'id = ' + convert(varchar(20), v_id_scet)
			);
        end if;
		if update(quantity) then
			call update_remote(remoteServerNew, 'scet', 'kol1', convert(varchar(20), v_quantity), 'id = ' + convert(varchar(20), v_id_scet));
		end if;
	end if;

end;
	
	
if exists (select 1 from systriggers where trigname = 'wf_delete_nomenk' and tname = 'sDmcRez') then 
	drop trigger sDmcRez.wf_delete_nomenk;
end if;
    
create TRIGGER wf_delete_nomenk before delete on
sDmcRez
referencing old as old_name
for each row
begin
	declare remoteServerNew varchar(32);
	declare v_id_jscet integer;
	
	select 
		sysname
		, id_jscet
	into 
		remoteServerNew
		, v_id_jscet
	from BayOrders o
	join GuideVenture v on o.ventureId = v.ventureId and v.standalone = 0
	where numOrder = old_name.numDoc;

	if remoteServerNew is not null then
		call delete_remote(remoteServerNew, 'scet', 'id = ' + convert(varchar(20), old_name.id_scet));
		call call_remote(remoteServerNew, 'slave_renu_scet', v_id_jscet);
	end if;
end;



----------------------------------------------------------------------
--------------         BayNomenkOut          -----------------
----------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_BayNomenkOut_outcome_di' and tname = 'BayNomenkOut') then 
	drop trigger BayNomenkOut.wf_BayNomenkOut_outcome_di;
end if;
/*
create 
	trigger wf_BayNomenkOut_outcome_di before delete order 1 on 
BayNomenkOut
referencing old as old_name
for each row
begin
--	declare v_id_mat integer;
	declare v_id_jmat integer;
	declare v_sysname varchar(50);

--	set v_id_mat = old_name.id_mat;
	set v_id_jmat = old_name.id_jmat;

	select v.sysname
	into v_sysname
	from BayOrders o
		left join guideventure v on v.ventureid = o.ventureid and v.standalone = 0
	where numorder = old_name.numOrder;

	call wf_otgruz_remove (
		v_id_jmat
		,'stime'
	);

	if v_sysname is not null and v_sysname != 'stime' then
		call wf_otgruz_remove (
			v_id_jmat
			,v_sysname
		);

	end if;

		
		
end;

*/



if exists (select 1 from systriggers where trigname = 'wf_BayNomenkOut_outcome_ui' and tname = 'BayNomenkOut') then 
	drop trigger BayNomenkOut.wf_BayNomenkOut_outcome_ui;
end if;

/*

create 
	trigger wf_BayNomenkOut_outcome_ui before update order 1 on 
BayNomenkOut
referencing new as new_name old as old_name
for each row
begin
	declare v_id_mat integer;
	declare v_id_jmat integer;
	declare v_sysname varchar(50);
	declare v_cena float;

	if update(quant) and old_name.quant != new_name.quant then
		set v_id_mat = old_name.id_mat;
		set v_id_jmat = old_name.id_jmat;

		select intQuant into v_cena from sDmcRez where numOrder = new_name.numOrder and nomnom = new_name.nomNom;

		select v.sysname
		into v_sysname
		from BayOrders o
		left join guideventure v on v.ventureid = o.ventureid and v.standalone = 0
		where numorder = old_name.numOrder;
		

		call wf_otgruz_quant(
			v_id_mat
			,v_id_jmat
			,new_name.quant
			,v_cena
			,'stime'
		);

		if v_sysname is not null and v_sysname != 'stime' then
			call wf_otgruz_quant(
				v_id_mat
				,v_id_jmat
				,new_name.quant
				,v_cena
				,v_sysname
			);

		end if;


	end if;
end;
*/


if exists (select 1 from systriggers where trigname = 'wf_BayNomenkOut_outcome_bi' and tname = 'BayNomenkOut') then 
	drop trigger BayNomenkOut.wf_BayNomenkOut_outcome_bi;
end if;
/*
create 
	trigger wf_BayNomenkOut_outcome_bi before insert order 1 on 
BayNomenkOut
referencing new as new_name
for each row
begin
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_jmat_nu varchar(20);
	declare v_mat_nu varchar(20);
	declare v_currency_rate float;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_id_source integer;
	declare v_id_dest integer;
--	declare v_osn varchar(100);
	declare v_id_jscet integer;
--	declare v_venture_id integer;
	declare v_firm_id integer;
	declare v_sysname varchar(50);
	declare v_ventureName varchar(100);
	declare v_cena float;
	declare v_cur_otgruz_date date;



	select max(id_jmat) into v_id_jmat 
	from BayNomenkOut 
	where numOrder = new_name.numOrder and outDate = new_name.outDate;

	select 
		 o.id_jscet
		, isnull(s.id_voc_names, 0)
		, isnull(f.id_voc_names,0)
		, v.ventureName
		, v.sysname
	into  
		 v_id_jscet
		, v_id_source
		, v_id_dest
		, v_ventureName
		, v_sysname
	from BayOrders o
		left join guideventure v on v.ventureid = o.ventureid and v.standalone = 0
		left join BayGuideFirms f on o.firmid = f.firmid
		left join sguidesource s on sourceid = -1001
	where numorder = new_name.numOrder;

	
	set v_id_currency = system_currency();
	call slave_currency_rate_stime(v_datev, v_currency_rate);

--	select id_voc_names into v_id_dest from guidefirms where firmid = v_firm_id;
--	    message 'v_id_dest = ', v_id_dest to client;
	-- со склада 1 
	-- ?? хотя по идее нужно бы отгружать со склада готовой продукции
--	select id_voc_names into v_id_source from sguidesource where sourceid = -1001;

	if v_id_jmat is null then
--	    message '---' to client;
		set v_id_jmat = wf_otgruz_jmat(
			new_name.numOrder
			, v_id_jscet
--			, v_venture_id
			, new_name.outDate
			, v_id_source
			, v_id_dest
			, v_id_currency
			, v_datev
			, v_currency_rate
			, v_sysname
		);
--		update BayOrders set id_jmat = v_id_jmat where numorder = new_name.numOrder;
	end if;

	set v_id_mat = get_nextid('mat');
	call slave_select_stime(v_mat_nu, 'mat', 'max(nu)', 'id_jmat = ' + convert(varchar(20), v_id_jmat));
	set v_mat_nu = convert(varchar(20), convert(integer, isnull(v_mat_nu, 0)) + 1);
	select intQuant into v_cena from sDmcRez where numDoc = new_name.numOrder and nomnom = new_name.nomNom;

	call wf_otgruz_nom(
		  v_id_mat
		, v_id_jmat
		, new_name.nomnom
		, new_name.quant
		, v_cena
		, v_mat_nu
		, v_id_source
		, v_id_dest
		, v_currency_rate
		, v_sysname
	);
	set new_name.id_mat = v_id_mat;
	set new_name.id_jmat = v_id_jmat;

end;

*/



----------------------------------------------------------------------
--------------         xUslugOut          -----------------
----------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_xUslugOut_outcome_di' and tname = 'xUslugOut') then 
	drop trigger xUslugOut.wf_xUslugOut_outcome_di;
end if;

/*
create 
	trigger wf_xUslugOut_outcome_di before delete order 1 on 
xUslugOut
referencing old as old_name
for each row
begin
--	declare v_id_mat integer;
	declare v_id_jmat integer;
	declare v_sysname varchar(50);

--	set v_id_mat = old_name.id_mat;
	set v_id_jmat = old_name.id_jmat;

	select v.sysname
	into v_sysname
	from Orders o
		left join guideventure v on v.ventureid = o.ventureid and v.standalone = 0
	where numorder = old_name.numOrder;

	call wf_otgruz_remove (
		v_id_jmat
		,'stime'
	);

	if v_sysname is not null and v_sysname != 'stime' then
		call wf_otgruz_remove (
			v_id_jmat
			,v_sysname
		);

	end if;

		
		
end;

*/



if exists (select 1 from systriggers where trigname = 'wf_xUslugOut_outcome_ui' and tname = 'xUslugOut') then 
	drop trigger xUslugOut.wf_xUslugOut_outcome_ui;
end if;

/*
create 
	trigger wf_xUslugOut_outcome_ui before update order 1 on 
xUslugOut
referencing new as new_name old as old_name
for each row
begin
	declare v_id_mat integer;
	declare v_id_jmat integer;
	declare v_sysname varchar(50);
	declare v_cena float;

	if update(quant) and old_name.quant != new_name.quant then
		set v_id_mat = old_name.id_mat;
		set v_id_jmat = old_name.id_jmat;

--		select intQuant into v_cena from sDmcRez where numOrder = new_name.numOrder and nomnom = new_name.nomNom;

		select v.sysname
			,o.ordered
		into v_sysname
			, v_cena
		from Orders o
		left join guideventure v on v.ventureid = o.ventureid and v.standalone = 0
		where numorder = old_name.numOrder;
		

		call wf_otgruz_quant(
			v_id_mat
			,v_id_jmat
			,new_name.quant
			,v_cena
			,'stime'
		);

		if v_sysname is not null and v_sysname != 'stime' then
			call wf_otgruz_quant(
				v_id_mat
				,v_id_jmat
				,new_name.quant
				,v_cena
				,v_sysname
			);

		end if;


	end if;
end;
*/




if exists (select 1 from systriggers where trigname = 'wf_xUslugOut_outcome_bi' and tname = 'xUslugOut') then 
	drop trigger xUslugOut.wf_xUslugOut_outcome_bi;
end if;

/*
create 
	trigger wf_xUslugOut_outcome_bi before insert order 1 on 
xUslugOut
referencing new as new_name
for each row
begin
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_jmat_nu varchar(20);
	declare v_mat_nu varchar(20);
	declare v_currency_rate float;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_id_source integer;
	declare v_id_dest integer;
--	declare v_osn varchar(100);
	declare v_id_jscet integer;
--	declare v_venture_id integer;
	declare v_firm_id integer;
	declare v_sysname varchar(50);
	declare v_ventureName varchar(100);
	declare v_cena float;
	declare v_cur_otgruz_date date;
	declare v_nomnom char(10);



	set v_nomnom = 'УСЛ';

	select max(id_jmat) into v_id_jmat 
	from xUslugOut 
	where numOrder = new_name.numOrder and outDate = new_name.outDate;

	select 
		 o.id_jscet
		, isnull(s.id_voc_names, 0)
		, isnull(f.id_voc_names,0)
		, v.ventureName
		, v.sysname
		, o.ordered
	into  
		 v_id_jscet
		, v_id_source
		, v_id_dest
		, v_ventureName
		, v_sysname
		, v_cena
	from Orders o
		left join guideventure v on v.ventureid = o.ventureid and v.standalone = 0
		left join BayGuideFirms f on o.firmid = f.firmid
		left join sguidesource s on sourceid = -1001
	where numorder = new_name.numOrder;

	
	set v_id_currency = system_currency();
	call slave_currency_rate_stime(v_datev, v_currency_rate);

--	select id_voc_names into v_id_dest from guidefirms where firmid = v_firm_id;
--	    message 'v_id_dest = ', v_id_dest to client;
	-- со склада 1 
	-- ?? хотя по идее нужно бы отгружать со склада готовой продукции
--	select id_voc_names into v_id_source from sguidesource where sourceid = -1001;

	if v_id_jmat is null then
--	    message '---' to client;
		set v_id_jmat = wf_otgruz_jmat(
			new_name.numOrder
			, v_id_jscet
--			, v_venture_id
			, new_name.outDate
			, v_id_source
			, v_id_dest
			, v_id_currency
			, v_datev
			, v_currency_rate
			, v_sysname
		);
	end if;

	set v_id_mat = get_nextid('mat');
	call slave_select_stime(v_mat_nu, 'mat', 'max(nu)', 'id_jmat = ' + convert(varchar(20), v_id_jmat));
	set v_mat_nu = convert(varchar(20), convert(integer, isnull(v_mat_nu, 0)) + 1);
--	select intQuant into v_cena from sDmcRez where numDoc = new_name.numOrder and nomnom = new_name.nomNom;

	call wf_otgruz_nom(
		  v_id_mat
		, v_id_jmat
		, v_nomnom
		, new_name.quant
		, v_cena
		, v_mat_nu
		, v_id_source
		, v_id_dest
		, v_currency_rate
		, v_sysname
	);
	set new_name.id_mat = v_id_mat;
	set new_name.id_jmat = v_id_jmat;

end;

*/



//===============================================
//    Процедуры обеспечения живучести программ
//===============================================

if exists (select 1 from sysprocedure where proc_name = 'get_standalone') then
	drop function get_standalone;
end if;



CREATE function get_standalone(
	 p_server varchar(50)
	 ,p_remote integer default 0
) returns integer
begin
	declare v_check varchar(23);

	if isnumeric(p_server)=1 then
		select standalone into v_check from guideVenture where ventureId = p_server;
	else
		select standalone into v_check from guideVenture where sysname = p_server;
	end if;
	if v_check is null then
		set get_standalone = 1;
	else 
		set get_standalone = v_check;
	end if;
end;



if exists (select 1 from sysprocedure where proc_name = 'slave_set_standalone') then
	drop function slave_set_standalone;
end if;

// return 1 - successful changing
//		  0 - failed

CREATE function slave_set_standalone(
	 p_status varchar(23)
	 ,p_server varchar(50) default null
	 ,p_remote integer default 0
) returns integer
begin
	set slave_set_standalone = 1;
	if isnumeric(p_server)=1 then
		update guideVenture set standalone = p_status where ventureId = p_server;
	else
		update guideVenture set standalone = p_status where sysname = p_server;
	end if;
	if p_remote = 1 and p_server is not null then
		execute immediate 'call slave_set_standalone_'+ p_server +'( slave_set_standalone, ''' + p_status + ''')';
//		call call_remote(p_server, 'slave_set_standalone', ''''+ p_status + '''');
	end if; 
	exception when others then
		set slave_set_standalone = 0;
end;



if exists (select 1 from sysprocedure where proc_name = 'get_standalone_remote') then
	drop function get_standalone_remote;
end if;

CREATE function get_standalone_remote(
	 p_server varchar(50) default null
) returns integer
begin
	set get_standalone_remote = 0;
	execute immediate 'call slave_get_standalone_'+ p_server +'( get_standalone_remote)';
	exception when others then
		set get_standalone_remote = -1;
end;


-------------------------------------------------------------------------
--------------             BayGuideFirms      ----------------------
-------------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_delete_firm' and tname = 'BayGuideFirms') then 
	drop trigger BayGuideFirms.wf_delete_firm;
end if;

create TRIGGER wf_delete_firm before delete on
BayGuideFirms
referencing old as old_name
for each row
begin
	if old_name.id_voc_names is not null then
		call delete_host('voc_names', 'id = ' + convert(varchar(20), old_name.id_voc_names));
	end if;
end;


if exists (select 1 from systriggers where trigname = 'wf_update_firm' and tname = 'BayGuideFirms') then 
	drop trigger BayGuideFirms.wf_update_firm;
end if;

create TRIGGER wf_update_firm before update on
BayGuideFirms
referencing old as old_name new as new_name
for each row
begin
	if update(phone) then
		call update_host('voc_names', 'phone', '''''' + new_name.phone + '''''', 'id = ' + convert(varchar(20), old_name.id_voc_names));
	end if;
	if update(fio) then 
		call update_host('voc_names', 'rem', '''''' + new_name.fio + '''''', 'id = ' + convert(varchar(20), old_name.id_voc_names));
	end if;
end;


if exists (select 1 from systriggers where trigname = 'wf_insert_firm' and tname = 'BayGuideFirms') then 
	drop trigger BayGuideFirms.wf_insert_firm;
end if;

create TRIGGER wf_insert_firm before insert on
BayGuideFirms
referencing new as new_name
for each row
begin
	declare v_zakaz_id integer;
	declare v_params varchar(2000);
	declare v_firms_id integer;

	select id_voc_names into v_zakaz_id from BayGuideFirms where firmid = 0;

	-- id  фирмы в базе Комтеха
	set v_firms_id = get_nextid ('voc_names');
	set v_params =
		 convert(varchar(20), v_firms_id)
		+ ', '''''+ substring(new_name.name,1,203) + ''''''
	;
	set v_params = v_params + ', ' + convert(varchar(20), v_zakaz_id);

	call insert_host('voc_names', 'id, nm, belong_id', v_params);

	set new_name.id_voc_names = v_firms_id;
	
end;

-------------------------------------------------------------------------
--------------             GuideFirms      ----------------------
-------------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_delete_firm' and tname = 'GuideFirms') then 
	drop trigger GuideFirms.wf_delete_firm;
end if;

create TRIGGER wf_delete_firm before delete on
GuideFirms
referencing old as old_name
for each row
begin
	if old_name.id_voc_names is not null then
		call delete_host('voc_names', 'id = ' + convert(varchar(20), old_name.id_voc_names));
	end if;
end;


if exists (select 1 from systriggers where trigname = 'wf_update_firm' and tname = 'GuideFirms') then 
	drop trigger GuideFirms.wf_update_firm;
end if;

create TRIGGER wf_update_firm before update on
GuideFirms
referencing old as old_name new as new_name
for each row
begin
	if update(phone) then
		call update_host('voc_names', 'phone', '''''' + new_name.phone + '''''', 'id = ' + convert(varchar(20), old_name.id_voc_names));
	end if;
	if update(fio) then 
		call update_host('voc_names', 'rem', '''''' + new_name.fio + '''''', 'id = ' + convert(varchar(20), old_name.id_voc_names));
	end if;
end;


if exists (select 1 from systriggers where trigname = 'wf_insert_firm' and tname = 'GuideFirms') then 
	drop trigger GuideFirms.wf_insert_firm;
end if;

create TRIGGER wf_insert_firm before insert on
GuideFirms
referencing new as new_name
for each row
begin
	declare v_zakaz_id integer;
	declare v_params varchar(2000);
	declare v_firms_id integer;

	select id_voc_names into v_zakaz_id from guidefirms where firmid = 0;

	-- id  фирмы в базе Комтеха
	set v_firms_id = get_nextid ('voc_names');
	set v_params =
		 convert(varchar(20), v_firms_id)
		+ ', '''''+ substring(new_name.name,1,203) + ''''''
	;
	set v_params = v_params + ', ' + convert(varchar(20), v_zakaz_id);

	call insert_host('voc_names', 'id, nm, belong_id', v_params);

	set new_name.id_voc_names = v_firms_id;
	
end;

-------------------------------------------------------------------------
--------------             sGuideSource      ----------------------
-------------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_delete_source' and tname = 'sGuideSource') then 
	drop trigger sGuideSource.wf_delete_source;
end if;

create TRIGGER wf_delete_source before delete on
sGuideSource
referencing old as old_name
for each row
begin
	if old_name.id_voc_names is not null then
		call delete_host('voc_names', 'id = ' + convert(varchar(20), old_name.id_voc_names));
	end if;
end;


if exists (select 1 from systriggers where trigname = 'wf_update_source' and tname = 'sGuideSource') then 
	drop trigger sGuideSource.wf_update_source;
end if;

create TRIGGER wf_update_source before update on
sGuideSource
referencing old as old_name new as new_name
for each row
begin
	if update(phone) then
		call update_host('voc_names', 'phone', '''''' + new_name.phone + '''''', 'id = ' + convert(varchar(20), old_name.id_voc_names));
	end if;
	if update(fio) then 
		call update_host('voc_names', 'rem', '''''' + new_name.fio + '''''', 'id = ' + convert(varchar(20), old_name.id_voc_names));
	end if;
end;


if exists (select 1 from systriggers where trigname = 'wf_insert_source' and tname = 'sGuideSource') then 
	drop trigger sGuideSource.wf_insert_source;
end if;

create TRIGGER wf_insert_source before insert on
sGuideSource
referencing new as new_name
for each row
begin
	declare v_postav_id integer;
	declare v_params varchar(2000);
	declare v_sources_id integer;

	if isnull(new_name.sourceId, 0) >= 0 then
		select id_voc_names into v_postav_id from sGuideSource where sourceid = 0;
	else 
		set v_postav_id = select_remote('stime', 'voc_names', 'id', 'belong_id = 0 and nm = ''''Объекты затрат''''');
	end if;

	-- id  фирмы в базе Комтеха
	set v_sources_id = get_nextid ('voc_names');
	set v_params =
		 convert(varchar(20), v_sources_id)
		+ ', '''''+ substring(new_name.sourceName,1,203) + ''''''
	;
	set v_params = v_params + ', ' + convert(varchar(20), v_postav_id);

	call insert_host('voc_names', 'id, nm, belong_id', v_params);

	set new_name.id_voc_names = v_sources_id;
	
end;




/*
----------------------------------------------------------------------
--------------                 sDocsVenture          ------------------------
----------------------------------------------------------------------
if exists (select 1 from systriggers where trigname = 'wf_delete' and tname = 'sDocsVenture') then 
	drop trigger sDocsVenture.wf_delete;
end if;

create TRIGGER wf_delete before delete on
sDocsVenture
referencing old as old_name
for each row
begin
	declare remoteServer varchar(32);
	declare no_echo integer;

	set no_echo = 0;

  	begin
		select @stime into no_echo; 
	exception 
		when other then
			set no_echo = 0;
	end;

	if no_echo = 1 then
		return;
	end if;



	if (old_name.id_jmat is not null) then
		call block_remote('stime', get_server_name(), 'jmat');
		call block_remote('stime', get_server_name(), 'mat');
		call delete_remote('stime', 'jmat', 'id = ' + convert(varchar(20), old_name.id_jmat));
		call unblock_remote('stime', get_server_name(), 'jmat');
		call unblock_remote('stime', get_server_name(), 'mat');
	end if;

end;



if exists (select 1 from systriggers where trigname = 'wf_insert' and tname = 'sDocsVenture') then 
	drop trigger sDocsVenture.wf_insert;
end if;

create 
	trigger wf_insert before insert on 
sDocsVenture
referencing new as new_name
for each row
begin
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_jmat_nu varchar(20);
	declare v_currency_rate float;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_id_source integer;
	declare v_id_dest integer;
	declare v_osn varchar(100);
	declare v_id_guide_jmat integer;




		set v_id_jmat = get_nextid('jmat');
		set v_id_currency = system_currency();
		call slave_currency_rate_stime(v_datev, v_currency_rate);
--		set v_jmat_nu = new_name.numdoc;
		select id_voc_names into v_id_source from sguidesource where sourceid = new_name.sourid;
		select id_voc_names into v_id_dest from sguidesource where sourceid = new_name.destid;
		set v_osn = '[Prior: '+ convert(varchar(20), new_name.numdoc) +']';
	    
		call wf_insert_jmat (
			'stime'
			,v_id_guide_jmat
			,v_id_jmat
			,now() --v_jmat_date
			,v_jmat_nu
			,v_osn
			,v_id_currency
			,v_datev
			,v_currency_rate
			,v_id_source
			,v_id_dest
		);
		set new_name.id_jmat = v_id_jmat;


end;

if exists (select 1 from systriggers where trigname = 'wf_update' and tname = 'sDocsVenture') then 
	drop trigger sDocsVenture.wf_update;
end if;

create 
	trigger wf_update before update on 
sDocsVenture
referencing new as new_name old as old_name
for each row
begin
	declare v_id_jmat integer;
	declare v_id_mat integer;
	declare v_jmat_nu varchar(20);
	declare v_currency_rate float;
	declare v_datev date;
	declare v_id_currency integer;
	declare v_id_source integer;
	declare v_id_dest integer;
	declare v_osn varchar(100);

	if update(sourid) then
		select id_voc_names into v_id_source from sguidesource where sourceid = new_name.sourid;
		call slave_update_stime('jmat', 'id_s', convert(varchar(20), v_id_source), 'id = ' + convert(varchar(20), old_name.id_jmat));
	end if;
	if update(destid) then
		select id_voc_names into v_id_dest from sguidesource where sourceid = new_name.destid;
		call slave_update_stime('jmat', 'id_d', convert(varchar(20), v_id_dest), 'id = ' + convert(varchar(20), old_name.id_jmat));
	end if;
	if update(note) then
		set v_osn = '[Prior: '+ new_name.note +']';
		call slave_update_stime('jmat', 'osn', '''' +v_osn + '''', 'id = ' + convert(varchar(20), old_name.id_jmat));
	end if;
end;
*/
