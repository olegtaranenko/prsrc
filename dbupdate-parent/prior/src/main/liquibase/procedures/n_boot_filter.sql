ALTER PROCEDURE "DBA"."n_boot_filter" (
	  p_filterid    integer
	, p_managId     varchar(16)
)
begin

--	insert into tmpNBootReport (paramName, paramValue)
	select p.name as paramName, ab.paramValue as paramValue
	from nAnalysBootingParam p 
	join nAnalysBooting ab on p.id = ab.paramId
	join nAnalys a on ab.templateId = a.templateId
	join nFilter f on f.byrowid = a.byrow and f.bycolumnid = a.bycolumn and f.id = p_filterId
	;
	
	-- 
--	select * from tmpNBootReport;

end