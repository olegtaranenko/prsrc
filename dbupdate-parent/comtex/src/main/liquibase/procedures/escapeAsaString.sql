if exists (select '*' from sysprocedure where proc_name like 'escapeAsaString') then  
	drop function escapeAsaString;
end if;

create function escapeAsaString (
	src long varchar
) returns long varchar
begin
	set escapeAsaString = replace(src, '''', '`');
end;
