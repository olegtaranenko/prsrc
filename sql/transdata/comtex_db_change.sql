
-- ��������� � id-��ࠧ������ �����, ����� ��諨 � ���ᨨ
-- ����� 8.1.5 (ﭢ��� 2006 ����).
-- ������ ⠡��� inc_table, � ������ ��宦� ᪫��뢠���� 
-- ����� id, ��� ⠡���� ����� ������ ���� �ᯮ�짮��� ��� 
-- ���������� ᫥���饩 ����� � ⠡����.
-- �஬� ��� ���������, ����� ���� ��ࠢ�� ⥪���� �����
-- ���ॡ���� �� ����� ��������� � ��楤��� ⨯� insert_host, 
-- insert_remote �, ����� ����, ��㣨� ��� ���४⭮�� ��ࠢ�����
-- ��� id.
-- TODO!!! ����� �㤥� ⠪�� �������� � 楫�� 㢥��祭�� 
-- �ந�����⥫쭮�� ��楤��� ����祭�� ᫥���饣� ������쭮�� id 
-- ��� ����客᪮� ⠡����

-- 31.10.2006 
--	��-�� ����: �訡�� �� ᬥ�� �।������ � ��室��� ���������
--	���ॡ������� �������� ������ ����祭��/䨪�஢���� nextid. ������ ���
--	����⢨⥫쭮 ������ �� inc_table.

begin
	declare nxt_id integer;

   	for d_cur as dc dynamic scroll cursor for
   		select table_nm as r_table_nm from inc_table
   	for update
	do
		execute immediate 'select max(id) into nxt_id from ' + r_table_nm;
		set nxt_id = isnull(nxt_id, 0) + 1;
		update inc_table set next_id = nxt_id where current of dc;

		call build_id_track_trigger(r_table_nm);
	end for;
end;



commit;
