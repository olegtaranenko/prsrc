if exists (select '*' from sysprocedure where proc_name like 'cmt_jscet_check') then  
	drop procedure cmt_jscet_check;
end if;

create procedure cmt_jscet_check (
	  in  p_id_jscet integer              // заказ, которому меняем номер счета
	, in  p_nu       varchar(10)          // новый номер счета заказа
	, out o_check    integer
)
begin

	declare v_curYear integer;

	select datepart(yy, dat) 
	into v_curYear
	from 
		jscet
	where id = p_id_jscet;

	select count(*) 
	into o_check
	from jscet js
	where js.nu = p_nu and datepart(yy, js.dat) = v_curYear and id != p_id_jscet;

end;

