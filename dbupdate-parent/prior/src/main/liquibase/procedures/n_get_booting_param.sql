ALTER FUNCTION "DBA"."n_get_booting_param" (
	  p_filterId integer
	, p_param_name varchar(64)
) returns varchar(127)
begin
	select paramValue
	into n_get_booting_param
	from nAnalysBooting ab
	join nAnalysBootingParam abp on abp.id = ab.paramId
	join nAnalys a on ab.templateId = a.templateId
	join nFilter f on f.byrowid = a.byrow and f.bycolumnid = a.bycolumn
	where abp.name = p_param_name
	and f.id = p_filterId;
end