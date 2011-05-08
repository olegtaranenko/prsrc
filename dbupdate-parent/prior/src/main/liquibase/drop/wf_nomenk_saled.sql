if exists (select '*' from sysprocedure where proc_name like 'wf_nomenk_saled') then  
	drop procedure wf_nomenk_saled;
end if;
