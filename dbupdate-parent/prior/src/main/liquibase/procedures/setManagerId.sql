if exists (select '*' from sysprocedure where proc_name like 'setManagerId') then  
	drop procedure setManagerId;
end if;

create procedure setManagerId (
	in p_manag    varchar(20)
)
begin
	declare v_managId tinyint;

    begin
		select  managId into v_managId
		from GuideManag where manag = p_manag;
		if v_managId is not null then
			set @managerId = v_managId;
		end if
    exception when others then
    	--set v_managId = null;
    end;
	
end;