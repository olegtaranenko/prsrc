if exists (select 1 from sysprocedure where proc_name = 'wf_dual_distribute') then
	drop procedure wf_dual_distribute;
end if;

create
	procedure wf_dual_distribute (
		  p_numdoc         integer
		, p_numext         integer
		, p_sourid         integer
		, p_destid         integer
		, p_dat            date
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
		,p_dat
		,v_jmat_nu
		,v_osn
		,v_id_currency
		,v_datev
		,v_currency_rate
		,v_id_source
		,v_id_dest
	);

end;


