if exists (select 1 from systriggers where trigname = 'wf_insert_gproduct' and tname = 'sguideproducts') then 
	drop trigger sguideproducts.wf_insert_gproduct;
end if;

create TRIGGER wf_insert_gproduct before insert on
sguideproducts
referencing new as new_name
for each row
begin
	declare v_id_inv integer;
	declare v_fields varchar(500);
	declare v_values varchar(2000);
	declare v_belong_id integer;
    declare v_id_edizm1 integer;
    declare v_id_size integer;
    declare v_nm varchar(102);


	set v_id_inv = get_nextid('inv');

	select id_inv into v_belong_id from sguideseries where seriaId = new_name.prSeriaId;
  	set v_id_edizm1 = wf_id_stuck();


  set v_fields = 
  	  ' id'
  	+ ',belong_id'
  	+ ',nomen'
    + ',nm'
    + ',prc1'
    + ',is_compl'
    + ', id_edizm1'
	;

	set v_nm = wf_make_invnm (new_name.prDescript, new_name.prSize, new_name.prName);

	set v_values = 
				 convert(varchar(20), v_id_inv)
		+ ', ' + convert(varchar(20), v_belong_id)
		+ ', ''''' + new_name.prName + ''''''
		+ ', ''''' + v_nm + ''''''
		+ ', ' + convert(varchar(20), new_name.cena4)
		+ ', 1'
   	  	+ ', '+convert(varchar(20), v_id_edizm1);
	;


	if new_name.prsize is not null and length(new_name.prsize) > 0 then
	  	set v_id_size = wf_getEdizmId(new_name.prsize);
   	  	set v_fields = v_fields + ', id_size';
   	  	set v_values = v_values + ', '+convert(varchar(20), v_id_size);
   	end if; 

	call insert_host('inv', v_fields, v_values);
  set new_name.id_inv=v_id_inv;
	

end;
