if exists (select '*' from sysprocedure where proc_name like 'wf_kreditor_debitor') then  
	drop function wf_kreditor_debitor;
end if;

create function wf_kreditor_debitor (
	p_id_deb integer
) returns integer
begin
	declare v_kredDebitor integer;
	declare varchar_id_deb varchar(20);
	declare v_values varchar(1024);
	declare v_deb_name varchar(203);

	
	select nm into v_deb_name from voc_names where id = p_id_deb;
	set v_deb_name = escapeAsaString(v_deb_name);

	call admin.slave_select_prior(varchar_id_deb, 'GuideFirms', 'min(firmid)', 'Name='''+ v_deb_name +'''');
	set v_kredDebitor = convert(integer, varchar_id_deb);

	if v_kredDebitor is null then
		call admin.slave_select_prior(varchar_id_deb, 'yDebKreditor', 'min(id)', 'Name='''+ v_deb_name +'''');
		set v_kredDebitor = convert(integer, varchar_id_deb);
	end if;

	if v_kredDebitor is null then
		call admin.slave_select_prior(varchar_id_deb, 'yDebKreditor', 'min(id)', '1=1');
		
		set v_kredDebitor = convert(integer, isnull(varchar_id_deb, '')) - 1;
		set v_values = '''' + convert(varchar(20), v_kredDebitor) + ''''
			+ ', ''' + v_deb_name + ''''
			+ ', '''  + admin.get_server_name() + ''''
		;
		call admin.slave_insert_prior('yDebKreditor', 'id, name, note', v_values);
	end if;

	return v_kredDebitor;
end;

