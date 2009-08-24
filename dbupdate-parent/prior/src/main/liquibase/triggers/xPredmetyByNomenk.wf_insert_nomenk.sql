if exists (select 1 from systriggers where trigname = 'wf_insert_nomenk' and tname = 'xPredmetyByNomenk') then 
	drop trigger xPredmetyByNomenk.wf_insert_nomenk;
end if;

create TRIGGER wf_insert_nomenk before insert on
xPredmetyByNomenk
referencing new as new_name
for each row
begin
	declare v_id_scet integer;
	declare v_id_jscet integer;
	declare v_id_inv integer;
	declare v_ventureid integer;
	declare remoteServerNew varchar(32);
	declare v_invCode varchar(10);
	declare v_fields varchar(255);
	declare v_values varchar(2000);
	declare scet_nu integer;
	declare v_date date;
	declare v_perList float;
	declare v_rate double;
	declare v_ndsrate       float;

	select id_jscet, ventureId, inDate, rate
	into v_id_jscet, v_ventureId, v_date, v_rate
	from orders 
	where numOrder = new_name.numOrder;

	select id_inv, perList 
	into v_id_inv, v_perList 
	from sGuideNomenk where nomNom = new_name.nomNom;

	select sysname, invCode 
		, v.nds
	into remoteServerNew, v_invcode 
		, v_ndsrate
	from GuideVenture v where ventureId = v_ventureId;

	if remoteServerNew is not null and v_id_jscet is not null then
	  -- Заказ, который имеет ссылки в бух.базах интеграции
	  -- т.е. уже назначен той, иди другой фирме
		set new_name.id_scet = 
			wf_insert_scet (
				remoteServerNew
				, v_id_jscet
				, v_id_inv
				, new_name.quant / v_perList
				, new_name.cenaEd
				, v_date
				, v_rate
				, v_ndsrate
			);
	end if;
	  
end;
