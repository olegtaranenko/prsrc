if exists (select '*' from sysprocedure where proc_name like 'put_jdog') then  
	drop procedure put_jdog;
end if;

create procedure put_jdog (
	  out o_id_jdog  integer
	, out o_nu_jdog  varchar(50)
	, in p_server    varchar(20)
	, in p_nu_jscet  varchar(50)
	, in p_id_post   integer
	, in p_dat       date
)
begin
	declare v_fields varchar(255);
	declare v_values varchar(2000);
	declare v_exists integer;

	set v_exists = select_remote(p_server, 'voc_names', 'count(*)', 'id = ' + convert(varchar(20), p_id_post));

	if v_exists = 1 then
	set o_nu_jdog = wf_make_jdog_nu(p_nu_jscet, now());
		set v_fields =
			'nu, id_post, dat, dat_end, dat_workbeg, dat_workend' -- , rem, nm
			;
    
		
		set v_values = 
				'''''' + o_nu_jdog + ''''''
			+ ', ' + convert(varchar(20), p_id_post)
			+ ', ''''' + convert(varchar(20), p_dat, 112) + ''''''
			+ ', ''''' + convert(varchar(20), p_dat, 112) + ''''''
			+ ', ''''' + convert(varchar(20), p_dat, 112) + ''''''
			+ ', ''''' + convert(varchar(20), p_dat, 112) + ''''''
		;
    
		set o_id_jdog = insert_count_remote(p_server, 'jdog', v_fields, v_values);
	end if;

end;

