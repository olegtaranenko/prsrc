
-- � ������� => ���
-- ����������� ������������� �������������� � ���������
call bootstrap_blocking();
--call inventory_order('20051013 21:00', 1, null);

-- ������������� ��� ���� �� �������
--call wf_cost_bulk_change(0);

-- � ������� => ���
-- �������������� �� ������������ (��, ��) �� ������� ����
--call v_inventory_order(null, '20041030 23:00');
--call v_inventory_order();

--����������� �������������, (� ������ ���������)
truncate table sdmcventure;
truncate table sdocsventure;

call ivo_generate(10);

commit;
exit;

begin
	for crs as c dynamic scroll cursor for
		select id as r_ivo_id from sdocsventure where cumulative_id is null
	do
		call ivo_to_comtex(r_ivo_id);
	end for;
end;
