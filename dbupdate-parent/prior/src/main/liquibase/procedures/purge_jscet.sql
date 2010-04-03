if exists (select '*' from sysprocedure where proc_name like 'purge_jscet') then  
	drop procedure purge_jscet;
end if;

create procedure purge_jscet (
		p_server varchar(32)
		, in p_id_jscet integer
)
begin
	call delete_remote(p_server, 'jscet', 'id = ' + convert(varchar(20), p_id_jscet));
	call delete_remote(p_server, 'scet', 'id_jmat = ' + convert(varchar(20), p_id_jscet));
	call delete_remote(p_server, 'jdog d', 'd.id != 0 and d.id = s.id_jdog and s.id = ' + convert(varchar(20), p_id_jscet), 'jscet s');
end;


