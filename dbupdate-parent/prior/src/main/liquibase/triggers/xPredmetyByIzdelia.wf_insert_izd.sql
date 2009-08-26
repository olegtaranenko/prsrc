if exists (select 1 from systriggers where trigname = 'wf_insert_izd' and tname = 'xPredmetyByIzdelia') then 
	drop trigger xPredmetyByIzdelia.wf_insert_izd;
end if;

create TRIGGER wf_insert_izd before insert on
xPredmetyByIzdelia
referencing new as new_name
for each row
begin
	declare v_id_scet       integer;
	declare v_id_jscet      integer;
	declare v_id_inv        integer;
--	declare v_ventureid integer;
	declare remoteServerNew varchar(32);
	declare v_invCode       varchar(10);
	declare v_fields        varchar(255);
	declare v_values        varchar(2000);
	declare v_date          date;
	declare v_rate          double;
	declare v_ndsrate       float;
 
	select id_jscet, inDate, sysname, invCode, o.rate
		, v.nds
	into v_id_jscet, v_date, remoteServerNew, v_invcode, v_rate
		, v_ndsrate
	from orders o
	join GuideVenture v on o.ventureId = v.ventureId and v.standalone = 0
	where numOrder = new_name.numOrder;

	select id_inv into v_id_inv 
		from sGuideProducts where prId = new_name.prId;
  
	if remoteServerNew is not null and v_id_jscet is not null then
		set v_id_scet =	
			wf_insert_scet (
				remoteServerNew
				, v_id_jscet
				, v_id_inv
				, new_name.quant
				, new_name.cenaEd
				, v_rate
				, v_ndsrate
			);
		set new_name.id_scet = v_id_scet;
		set new_name.id_inv = v_id_inv;
	end if;
end;

