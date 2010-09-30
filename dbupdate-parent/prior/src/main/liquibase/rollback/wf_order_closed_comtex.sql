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
		-- ��� ��������� - �� ������ �������� �� ��������.
		return;
	end if;

	select tp into v_orders_table from all_orders where numorder = p_numorder;

	execute immediate 'select id_jscet into v_old_id_jscet '
		+ 'from '+ v_orders_table + ' where numorder = ' + convert(varchar(20), p_numorder);

	if      v_old_id_jscet is not null
		and p_sysname != 'stime' -- ������ ��� �� � ��
	then
		-- ��������� ������ �� ����� � �����������
		set v_gad_level = select_remote(p_sysname, 'jscet', 'data_lock', 
			'id = ' + convert(varchar(20), v_old_id_jscet)
		);
		if v_gad_level = 0 then
			set wf_order_closed_comtex = 0;
			--raiserror 17001 '������ ������� �����, �� ��� ���, ���� �� �� ������ � �����������';
		end if;
	end if;

end;

