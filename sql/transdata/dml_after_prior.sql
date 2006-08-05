/*

begin
	declare v_nm varchar(150);
	for aCursor as a dynamic scroll cursor for
		select 
			  c.xprext as r_xprext
			, c.id_inv as r_id_inv_variant
			, p.prName as r_prName
			, p.prSize as  r_prSize
			, p.prDescript as r_prDescript
		from sguidecomplect c
		join sguideproducts p on p.prid = c.productid
	do
		set v_nm = wf_make_variant_nm (
			  r_prDescript
			, r_prSize
			, r_prName
			, r_xprext
		);
		call update_host('inv', 'nm', '''''' + v_nm + '''''', 'id = ' + convert(varchar(20), r_id_inv_variant));
		call update_host('inv', 'nomen', '''''' + r_prName + '''''', 'id = ' + convert(varchar(20), r_id_inv_variant));

		end for;
end;
*/


/*

call bootstrap_blocking;




-- Если бухгалтер приход вобьет в stime, то поможет этот скрипт.
-- он перенесет эти приходы в комтеховскую аналитику.

begin
	declare v_id_mat integer;
	declare v_mat_nu integer;
	declare v_currency_rate float;
	declare v_cost float;
	declare v_datev date;
	
	call slave_currency_rate_stime(v_datev, v_currency_rate);
	
	set v_cost = 0;
	set v_currency_rate = 30;
	set v_mat_nu = 1;
	
	call block_remote('stime', @@servername, 'mat');

	for cur as n dynamic scroll cursor for
		select 
			id_jmat as r_id_jmat
			, k.id_Inv as r_id_inv
			, s.id_voc_names as r_id_src
			, d.id_voc_names as r_id_dst
			, k.perlist as r_perlist
			, m.quant as r_qty
			, m.numdoc as r_numdoc
			, m.numext as r_numext
			, m.nomnom as r_nomnom
			, k.cena1 as r_cost
		from sdocs c
		join sdmc m on m.numdoc = c.numdoc and m.numext = c.numext and m.id_mat is null
		join sguidesource s on s.sourceid = c.sourid
		join sguidesource d on d.sourceid = c.destid
		join sguidenomenk k on k.nomnom = m.nomnom
		where id_jmat is not null and c.numext = 255
	do
			set v_id_mat = wf_insert_mat (
				'stime'
				,null
				,r_id_jmat
				,r_id_inv
				,v_mat_nu
				,r_qty 
				,r_cost
				,v_currency_rate
				,r_id_src
				,r_id_dst
				,r_perList
			);
			
		update sDmc set id_mat = v_id_mat where numDoc = r_numdoc and numext = r_numext and nomnom = r_nomnom;
		set v_mat_nu = v_mat_nu + 1;
	end for;
	call unblock_remote('stime', @@servername, 'mat');

end;
*/

call bootstrap_blocking;

begin 
	declare v_procent float;
	declare v_date date;
	declare v_date_end date;
	declare v_nomnom varchar(20);
--	set v_nomnom = '1002И6110';

	select ivo_procent into v_procent from system;
--	select activity_start into v_date from guideventure where sysname = 'markmaster';
--	set v_date_end = now();

	delete from sdocsventure;

	call fill_venture_order (
		  v_procent
		, v_date
		, v_date_end
		, v_nomnom
	);
end;



commit;


