-- ������� remote-�������. 
-- ��� ����� ������ ������ ��� �������� legacy ������

/*
call build_remote_table('inv', 1);
call build_remote_table('jmat', 1);
call build_remote_table('mat', 1);
call build_remote_table('jscet', 1);
call build_remote_table('scet', 1);
*/
-- ����� ��� ������ ���������� � ������-������������
-- � ����� FirmComtex
-- TODO ���� �� ������� ������� ODBC, ������� 
-- ���� �� �������� ������ �� ������-���������
-- �� ���������� � ���� �������� ������ ��
--call build_remote_table('voc_names', 1);
--call build_remote_table('post', 1);
--call build_remote_table('jmat', 1);
--call build_remote_table('mat', 1);

call build_table_one_server('jmat', 'stime', 1);
call build_table_one_server('mat', 'stime', 1);

call build_remote_table('xoz', 1);
