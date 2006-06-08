-- изменени€ к 18 окт€бр€ 2005 года
/*
begin
	declare nindex integer;
	set nindex = 0;
	for aCursor as c_nom dynamic scroll cursor for
		select 
			  e1.id_edizm as r_id_edizm1
			, e2.id_edizm as r_id_edizm2
			, n.cost as prc1
			, id_inv as r_id_inv
			, perList as r_perlist
		from sguidenomenk n
			left join edizm e1 on n.ed_izmer = e1.name
			left join edizm e2 on n.ed_izmer2 = e2.name
		where n.perlist != 1
	do
		call update_host('inv', 'id_edizm1', convert(varchar(20), r_id_edizm2), 'id = ' + convert(varchar(20), r_id_inv));
		call update_host('inv', 'id_edizm2', convert(varchar(20), r_id_edizm1), 'id = ' + convert(varchar(20), r_id_inv));


	end for;

end;
*/

-- изменени€ к 19 окт€бр€ 2005 года
begin
	declare nindex integer;
	declare v_currency_rate float;

	set v_currency_rate = 30.0;
	set nindex = 0;
	for aCursor as c_nom dynamic scroll cursor for
		select 
			 n.cost as r_prc1
			, id_inv as r_id_inv
			, perList as r_perlist
		from sguidenomenk n
	do
		call update_host('inv', 'prc1', convert(varchar(20), round(v_currency_rate * r_prc1, 2)), 'id = ' + convert(varchar(20), r_id_inv));
	end for;


	for aCursor as c_izd dynamic scroll cursor for
		select 
			 id_inv as r_id_inv
			, n.cena4 as r_prc1
		from sguideproducts n
	DO

		call update_host('inv', 'prc1', convert(varchar(20), round(v_currency_rate * r_prc1, 2)), 'id = ' + convert(varchar(20), r_id_inv));

	end for;

end;
commit;