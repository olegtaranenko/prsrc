if exists (select '*' from sysprocedure where proc_name like 'wf_make_prnm') then  
	drop function wf_make_prnm;
end if;


create 
 	function wf_make_prnm (
	  p_prName varchar(50) default null
	, p_ext integer default null
) returns varchar(150)
begin
	set wf_make_prnm = '';
	if isnull(p_ext, 0) <> 0 then
    	set wf_make_prnm = convert(varchar(8), p_ext) + '/';
	end if;
	set wf_make_prnm = convert(varchar(20), wf_make_prnm) + p_prName;

end;

