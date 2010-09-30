if exists (select '*' from sysprocedure where proc_name like 'wf_report_bright_rpf') then  
	drop procedure wf_report_bright_rpf;
end if;

create procedure wf_report_bright_rpf (
	p_prodCategoryId integer
--	, p_csvFlag integer default 0
)
begin

	declare v_ord_table varchar(64);
	declare p_table_name varchar(64);    
	declare p_id_name varchar(64);       
	declare p_parent_id_name varchar(64);
	declare p_order_by_name varchar(256);
	set p_table_name = 'sGuideSeries';
	set v_ord_table = get_tmp_ord_table_name(p_table_name);
	execute immediate get_tmp_ord_create_sql(v_ord_table);

	set p_id_name = 'seriaId';
	set p_parent_id_name = 'parentSeriaId';
	set p_order_by_name = 'seriaName';
	call wf_sort_klassificator(p_table_name, p_id_name, p_parent_id_name, p_order_by_name);


	select 
		  ph.prId, ph.prName, ph.prSeriaId, ph.prSize, ph.prDescript 
		, ph.vremObr, ph.formulaNom, ph.cena4, ph.page, ph.sortNom
		, ph.rabbat as productRabbat
		, s.gain2, s.gain3, s.gain4
		, f.formula, w.prId as hasWeb
		--, p.quantity as quantEd
	from sGuideProducts ph
	join #sGuideSeries_ord os on os.id       = ph.prSeriaId
	left join wf_izdeliaWithWeb w on w.prId  = ph.prId
	join sGuideSeries       s on s.seriaId   = ph.prSeriaId
	left join sGuideFormuls f on f.nomer     = ph.formulaNom
	where 
			ph.prodCategoryId = 2 
		and isnumeric(ph.page) = 1
		and isnull(p_prId, ph.prId) = ph.prId 
	order by os.ord, ph.sortNom;


	drop table #sGuideSeries_ord;

/*

	EXECUTE IMMEDIATE WITH RESULT SET ON не работает в asa 8.0, только в 9-ке :(

	declare v_sql long varchar;
	if p_csvFlag = 0 then
		set p_table_name = 'sGuideSeries';
		set v_ord_table = get_tmp_ord_table_name(p_table_name);
		execute immediate get_tmp_ord_create_sql(v_ord_table);
	    
		set p_id_name = 'seriaId';
		set p_parent_id_name = 'parentSeriaId';
		set p_order_by_name = 'seriaName';
		call wf_sort_klassificator(p_table_name, p_id_name, p_parent_id_name, p_order_by_name);
	else 
	end if;


	set v_sql = 
		'	select '
		+ '		  ph.prId, ph.prName, ph.prSeriaId, ph.prSize, ph.prDescript '
		+ '		, ph.vremObr, ph.formulaNom, ph.cena4, ph.page, ph.sortNom   '
		+ '		, ph.rabbat as productRabbat                                 '
		+ '		, s.gain2, s.gain3, s.gain4                                  '
		+ '		, f.formula'
	;
	if p_csvFlag = 0 then
		set v_sql = v_sql
		+ '		, w.prId as hasWeb                                '
	end if;

	set v_sql = v_sql
        + '	from sGuideProducts ph                                           '
		+ '	join sGuideSeries       s on s.seriaId   = ph.prSeriaId          '
        + '	left join sGuideFormuls f on f.nomer     = ph.formulaNom         '
    ;
	if p_csvFlag = 0 then
		set v_sql = v_sql
		+ '	join #sGuideSeries_ord os on os.id       = ph.prSeriaId          '
		+ '	left join wf_izdeliaWithWeb w on w.prId  = ph.prId               '
	end if;

	set v_sql = v_sql
		+ '	where                                                            '
		+ '			ph.prodCategoryId =  p_prodCategoryId'
		+ '		and isnumeric(ph.page) = 1                                   '
        + '	order by '
	;
	if p_csvFlag = 0 then
		set v_sql = v_sql
        + '	os.ord, ph.sortNom                                     '
	else 
		set v_sql = v_sql
        + '	ph.cod                                     '
	end if;

	EXECUTE IMMEDIATE WITH RESULT SET ON v_sql;
*/

end;
