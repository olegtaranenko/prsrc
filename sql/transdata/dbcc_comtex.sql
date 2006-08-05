
/* ���������, ��������� � id ���������� (������� inc_table)
���������� � comtex_db_change.sql, ��� ��� ��� �� ���� ������������ ��������� 
���� � ��� ��������� update ������ ��� ������ ����� ��������� �����������
*/


-- ��������� ��� ����������� ���� �����, ��� ����� � ����������� ������,
delete from guides_access_data where guide_id = 1005;

begin
	declare v_table_name varchar(128);
	declare v_column_name varchar(128);
	declare v_status_close_id integer;
	declare v_trigger_sql varchar(3000);
	declare v_sql varchar(1000);
	

	-- ����� ���������������� ���������� � ������� � ������� �������
	-- �������� ���-�� ���� ����� 'GUIDE_803_129574.NM','JSCET__USER_129573'
	select nm, parent_col_name
	into v_table_name, v_column_name
	from browsers where id_guides = 1005 
	and nm like '%guid%' 
	and namer like '%���%';

	if v_table_name is null then 
		return;
	end if;
	-- ������� ��  'GUIDE_803','USER_129573'
	set v_table_name = 'GUIDE_' + substring(v_table_name, 7, charindex('_', substring(v_table_name, 7))-1);
	set v_column_name =  substring(v_column_name, charindex('__', v_column_name)+2);
	-- 
	execute immediate 'select id into v_status_close_id from ' + v_table_name + ' where nm = ''��''';

	set v_sql = 'insert into guides_access_data  (guide_id, data_id, access_level)'
		+ '\nselect 1005, s.id, 1'
		+ '\nfrom jscet s '
		+ '\nwhere ' + v_column_name + ' = '+ convert(varchar(20), v_status_close_id)
		+ '\nand dat > convert(date, ''20051013'')';

	execute immediate v_sql;

end;



commit;
