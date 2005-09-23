begin
declare icount int;
declare folder_id integer;
declare v_id_cur integer;

select count(*) into icount from size;
if icount <= 0 then
	call legacy_guides();
end if;


--���������������� ������
select id into folder_id from voc_names_st where nm = '������' and belong_id = 0;
call move_old_voc_names(folder_id);
call legacy_sklad();


--������� ������
select id into folder_id from voc_names_st where nm = '������� ������' and belong_id = 0;
call move_old_voc_names(folder_id);
call legacy_zatr();


--��������� �����������
select id into folder_id from voc_names_st where nm = '��������� �����������' and belong_id = 0;
call move_old_voc_names(folder_id);
call legacy_firms();

-- ��������� ������ "�������� ������" � ���� ��
call legacy_currency();


-- ������������
select id into folder_id from inv_st where nm = '���������' and belong_id = 0;
call move_old_inv(folder_id);
select id into folder_id from inv_st where nm = '�������' and belong_id = 0;
call move_old_inv(folder_id);
call legacy_inv();


-- ���������� ��������������� ������ ��� ���������� �������
call host_legacy_variant();

-- ��������� ��������� ������ �� ������ � ������������� ���� st
call legacy_income_order();


update system set trans_date = now();

commit;

end;
