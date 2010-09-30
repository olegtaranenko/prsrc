if exists (select '*' from sysprocedure where proc_name like 'wf_report_price_web_izdelia') then  
	drop procedure wf_report_price_web_izdelia;
end if;

create procedure wf_report_price_web_izdelia (
)
begin
	
	declare v_ord_table varchar(64);

end;