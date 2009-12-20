if exists (select 1 from systriggers where trigname = 'wf_update_gproducts' and tname = 'sGuideProducts') then 
	drop trigger sGuideProducts.wf_update_gproducts;
end if;

create TRIGGER wf_update_gproducts before update on
sGuideProducts
referencing old as old_name new as new_name
for each row
begin
	declare v_id_inv integer;
    declare v_belong_id integer;
    declare v_id_edizm integer;
    declare v_id_size integer;
    declare v_prDescript varchar(50);
    declare v_prSize varchar(30);
    declare v_prName varchar(20);
    declare v_nm varchar(102);
    declare is_variant integer;

	declare v_values varchar(100);
	declare v_fields varchar(200);
	
	set v_id_inv = old_name.id_inv;


  if update(prSize) or update(prName) or update (prDescript) then

	select 1 into is_variant from svariantpower vp where vp.productid = old_name.prId;

	if (new_name.prDescript != old_name.prDescript) then
		set v_prDescript = new_name.prDescript;
	else 
		set v_prDescript = old_name.prDescript;
	end if;
  
	if (new_name.prName != old_name.prName) then
		set v_prName = new_name.prName;
		call update_host('inv', 'nomen', '''''' + new_name.prName + '''''', 'id = ' + convert(varchar(20), v_id_inv));
		if is_variant is not null then
			call update_host('inv', 'nomen', '''''' + new_name.prName + '''''', 'belong_id = ' + convert(varchar(20), v_id_inv));
		end if;
	else 
		set v_prName = old_name.prName;
	end if;
  
	if (new_name.prSize != old_name.prSize) then
		set v_prSize = new_name.prSize;
		set v_id_size = wf_getSizeId(v_prSize);
		call update_host('inv', 'id_size', convert(varchar(20), v_id_size), 'id = ' + convert(varchar(20), v_id_inv));
		if is_variant is not null then
			call update_host('inv', 'id_size', convert(varchar(20), v_id_size), 'belong_id = ' + convert(varchar(20), v_id_inv));
		end if;
	else 
		set v_prSize = old_name.prSize;
	end if;
  
  
	set v_nm = wf_make_invnm (v_prDescript, v_prSize, v_prName);
	call update_host('inv', 'nm', '''''' + v_nm + '''''', 'id = ' + convert(varchar(20), v_id_inv));
	if is_variant is not null then
		
		for aCursor as a dynamic scroll cursor for
			select 
				  xprext as r_xprext
				, id_inv as r_id_inv_variant
			from sguidecomplect g
			where productid = old_name.prid
		do
			set v_nm = wf_make_variant_nm (
				  v_prDescript
				, v_prSize
				, v_prName
				, r_xprext
			);
			call update_host('inv', 'nm', '''''' + v_nm + '''''', 'id = ' + convert(varchar(20), r_id_inv_variant));
			call update_host('inv', 'nomen', '''''' + v_prName + '''''', 'id = ' + convert(varchar(20), r_id_inv_variant));

		end for;
	end if;

  end if;

/*
  if update(prName) then
	call update_host('inv', 'nomen', '''''' + new_name.prName + '''''', 'id = ' + convert(varchar(20), v_id_inv));
  end if;

  if update(prDescript) then
	call update_host('inv', 'nm', '''''' + new_name.prDescript + '''''', 'id = ' + convert(varchar(20), v_id_inv));
  end if;

  if update(prsize) then
  	set v_id_size = wf_getSizeId(new_name.prsize);
		call update_host('inv', 'id_size', convert(varchar(20), v_id_size), 'id = ' + convert(varchar(20), v_id_inv));
  end if;
*/


  if update(seriaId) then
	select id_inv into v_belong_id from sguideseries where seriaId = new_name.prSeriaId;
	call update_host('inv', 'belong_id', convert(varchar(20), v_belong_id), 'id = ' + convert(varchar(20), v_id_inv));
  end if;
  
end;

