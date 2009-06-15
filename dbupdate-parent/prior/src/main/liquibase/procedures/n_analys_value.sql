if exists (select 1 from sysprocedure where proc_name = 'n_analys_value') then
	drop function n_analys_value;
end if;

--	Выкатить наружу одиночное значение загрузочного параметра для анализа
	

CREATE function n_analys_value (
	  p_analysid    integer
	, p_analys_key  varchar(64)
) returns varchar (512)

begin

	select ab.paramValue 
		into n_analys_value
	from nAnalysBooting ab 
		join nAnalysBootingParam p on p.id = ab.paramId and p.name = p_analys_key
		join nAnalys a on a.id = p_analysid and ab.templateId = a.templateid
	;
	
end;

