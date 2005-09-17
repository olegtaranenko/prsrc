/*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/
/*                                                   */
/*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*/
if exists (select 1 from sysprocedure where proc_name = 'check_integrity_setting') then
	drop procedure check_integrity_setting;
end if;

CREATE procedure check_integrity_setting()
begin
end;
