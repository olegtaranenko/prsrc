-- ����������� � ������. �����-�� ������ ������ ������ ��������
-- ��������� �� � ���� rem
begin 
	for aCursor as b1 dynamic scroll cursor for
		select 
			f.firmId as f_id
			,f.name as r_name
			,f.fio as f_fio
			,f.phone as f_phone
			,f.email as f_email
			,id_voc_names as r_id_voc_names
		from guidefirms f
	DO
		call update_host ('voc_names', 'rem', 'address', 'id = ' + convert(varchar(20), r_id_voc_names));
		call update_host ('voc_names', 'address', '''''''''', 'id = ' + convert(varchar(20), r_id_voc_names));
	end for;
end;

begin 
	declare v_rem varchar(100);
	for aCursor as b1 dynamic scroll cursor for
		select 
			f.firmId as f_id
			,f.name as r_name
			,f.fio as f_fio
			,f.phone as f_phone
			,f.email as f_email
			,id_voc_names as r_id_voc_names
		from bayguidefirms f
	DO
		set v_rem = select_remote('stime', 'voc_names', 'rem', 'id = ' + convert(varchar(20), r_id_voc_names));
		if v_rem is null or char_length(v_rem) = 0 then
			call update_host ('voc_names', 'rem', 'address', 'id = ' + convert(varchar(20), r_id_voc_names));
			call update_host ('voc_names', 'address', '''''''''', 'id = ' + convert(varchar(20), r_id_voc_names));
		end if;
	end for;
end;



commit;

/*
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

call legacy_purpose();

update system set trans_date = now();

commit;

end;
*/