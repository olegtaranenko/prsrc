if exists (select 1 from sysprocedure where proc_name = 'temp_fedokruga') then
	drop procedure temp_fedokruga;
end if;

CREATE procedure temp_fedokruga(
)
begin
	declare v_regionId integer;
	declare v_territoryId integer;
	declare v_space_pos integer;
	declare v_partName varchar(64);
	declare v_nameFO varchar(64);

	-- две омские области
	update bayguidefirms set regionid = 78 where regionid = 41;
	delete from bayregion where regionid = 41;
	--Таймырский (Долгано - Нен, Усть - Ордынский Бурятски, Агинский Бурятский АО, Корякский АО не существует
	delete from bayregion where regionid in (85, 87, 89, 90);




	set v_territoryId = 0;

	for x as xc dynamic scroll cursor for
		select nameFO as r_nameFO, id as r_id from tempFO order by id
		for update
	do
		if substring (r_nameFO, 1, 3) = '   ' then
			set v_territoryId = r_id;
			set v_nameFO = substring(r_nameFO, 4);
			update tempFO set nameFO = v_nameFO, parentId = null where current of xc;
		else
			set v_nameFO = r_nameFO;
			update tempFO set parentId = v_territoryId where current of xc;
		end if;

		-- искать аналогичную в bayRegion
		set v_space_pos = charindex(' ', v_nameFO);
		if v_space_pos > 1 then
			set v_partName = substring(v_nameFO, 1, v_space_pos - 1);
			if v_partName = 'Республика' then
				set v_partName = substring (v_nameFO, v_space_pos + 1);
			elseif substring(v_nameFO, 2, 1) = '.' then
				set v_partName = substring (v_nameFO, v_space_pos + 1);
				set v_space_pos = charindex(' ', v_partName);
				set v_partName = substring (v_partName, v_space_pos - 1);
			end if;

		end if;

		set v_regionId = 0;
		select max(regionId) into v_regionId from BayRegion where Region like '%' + v_partName + '%';

		if isnull(v_regionId, 0) != 0 then
			update tempFO set BayRegionId = v_regionId where current of xc;
		end if;

	end for;

	update tempFO 
	set tempFO.bayRegionId = r.regionId 
	from bayRegion r 
	where r.region = tempFO.nameBayRegion and tempFO.nameBayRegion != '';

	insert into bayRegion (region) 
	select nameFO from tempFO fo where fo.bayregionId is null;


update bayregion set region = fo.namefo
from tempFO fo 
where fo.bayregionid = bayregion.regionid;


update tempFO set BayRegionid = r.regionId 
from bayregion r 
where tempFO.bayregionId is null and tempFO.nameFO = r.region;


update (bayregion 
join tempFO fo on bayregion.regionid  = fo.bayregionid
join bayregion r on r.regionid = fo.bayregionid
left join tempFO pfo on pfo.id = fo.parentid)
set bayregion.territoryid = pfo.bayregionid
where pfo.bayregionid is not null;
	
update bayregion set territoryid = fo.territoryid 
from tempFO fo
where fo.bayregionid = bayregion.regionid and fo.territoryid != 0;

end;
