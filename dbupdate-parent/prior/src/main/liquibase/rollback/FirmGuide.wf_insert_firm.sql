if exists (select 1 from systriggers where trigname = 'wf_insert_firm' and tname = 'FirmGuide') then 
	drop trigger FirmGuide.wf_insert_firm;
end if;


create TRIGGER wf_insert_firm before insert on
FirmGuide
referencing new as new_name
for each row
begin
	declare v_zakaz_id integer;
	declare v_params varchar(2000);
	declare v_firms_id integer;

	select id_voc_names into v_zakaz_id from FirmGuide where firmid = 0;

	-- id  фирмы в базе Комтеха
	set v_firms_id = get_nextid ('voc_names');
	set v_params =
		 convert(varchar(20), v_firms_id)
		+ ', '''''+ substring(new_name.name,1,203) + ''''''
	;
	set v_params = v_params + ', ' + convert(varchar(20), v_zakaz_id);

	call insert_host('voc_names', 'id, nm, belong_id', v_params);

	set new_name.id_voc_names = v_firms_id;
	
end;

