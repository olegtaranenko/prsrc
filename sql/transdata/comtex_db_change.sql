
-- ��������� � id-����������� �������, ������� ������ � ������
-- ������ 8.1.5 (������ 2006 ����).
-- ��������� ������� inc_table, � ������� ������ ������������ 
-- ����� id, ��� ������� ������� ������ ���� ����������� ��� 
-- ���������� ��������� ������ � �������.
-- ����� ���� ���������, ������� ������ �������� ������� ��������
-- ����������� ��� ������ ��������� � ��������� ���� insert_host, 
-- insert_remote �, ����� ����, ������ ��� ����������� �����������
-- ���� id.
-- TODO!!! ����� ����� ����� �������� � ����� ���������� 
-- ������������������ ��������� ��������� ���������� ����������� id 
-- ��� ������������ �������

-- 31.10.2006 
--	��-�� ����: ������ ��� ����� ������������ � ��������� ���������
--	������������� �������� ������ ���������/������������ nextid. ������ ���
--	������������� ������� �� inc_table.

/*
begin
	declare nxt_id integer;

   	for d_cur as dc dynamic scroll cursor for
   		select table_nm as r_table_nm from inc_table
   	for update
	do
		execute immediate 'select max(id) into nxt_id from ' + r_table_nm;
		set nxt_id = isnull(nxt_id, 1);
		update inc_table set next_id = nxt_id where current of dc;

		--call build_id_track_trigger(r_table_nm);
		call drop_id_track_trigger(r_table_nm);
	end for;
end;
*/


commit;
