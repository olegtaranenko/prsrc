/*
if exists (select '*' from sysprocedure where proc_name like 'slave_nextid_stime') then  
	drop procedure slave_nextid_stime;
end if;
if exists (select '*' from sysprocedure where proc_name like 'slave_count_delete_pm') then  
	drop procedure slave_nextid_pm;
end if;
if exists (select '*' from sysprocedure where proc_name like 'slave_count_delete_mm') then  
	drop procedure slave_nextid_mm;
end if;

	create PROCEDURE slave_nextid_stime(
			IN table_name char(100)
			, out id int
	)
	at 'stime;;ADMIN;slave_nextid';


	create PROCEDURE slave_nextid_pm(
	IN table_name char(100), out id int
	)
	at 'pm;;ADMIN;slave_nextid';

	create PROCEDURE slave_nextid_mm(
	IN table_name char(100), out id int
	)
	at 'mm;;ADMIN;slave_nextid';




--****************************************************************
--                               DELETE
--****************************************************************

if exists (select '*' from sysprocedure where proc_name like 'slave_delete_stime') then  
	drop procedure slave_delete_stime;
end if;
if exists (select '*' from sysprocedure where proc_name like 'slave_delete_mm') then  
	drop procedure slave_delete_mm;
end if;
if exists (select '*' from sysprocedure where proc_name like 'slave_delete_pm') then  
	drop procedure slave_delete_pm;
end if;

	create PROCEDURE "slave_delete_stime"(in table_name char(50),in where_cond char(2000))
	at 'stime;;ADMIN;slave_delete';

	create PROCEDURE "slave_delete_pm"(in table_name char(50),in where_cond char(2000))
	at 'pm;;ADMIN;slave_delete';

	create PROCEDURE "slave_delete_mm"(
			in table_name char(50)
			,in where_cond char(2000)
			)
	at 'mm;;ADMIN;slave_delete';




if exists (select '*' from sysprocedure where proc_name like 'slave_count_delete_stime') then  
	drop procedure slave_count_delete_stime;
end if;
if exists (select '*' from sysprocedure where proc_name like 'slave_count_delete_mm') then  
	drop procedure slave_count_delete_mm;
end if;
if exists (select '*' from sysprocedure where proc_name like 'slave_count_delete_pm') then  
	drop procedure slave_count_delete_pm;
end if;

	create PROCEDURE slave_count_delete_stime(
			out deleted integer
			,in table_name char(50)
			,in where_cond char(2000)
	)
	at 'stime;;ADMIN;slave_count_delete';

	create PROCEDURE slave_count_delete_pm(
			out deleted integer
			,in table_name char(50)
			,in where_cond char(2000)
	)
	at 'pm;;ADMIN;slave_count_delete';

	create PROCEDURE slave_count_delete_mm(
			out deleted integer
			,in table_name char(50)
			,in where_cond char(2000)
	)
	at 'mm;;ADMIN;slave_count_delete';



****************************************************************
                               INSERT
********************************************************************
if exists (select '*' from sysprocedure where proc_name like 'slave_insert_stime') then  
	drop procedure slave_insert_stime;
end if;
if exists (select '*' from sysprocedure where proc_name like 'slave_insert_pm') then  
	drop procedure slave_insert_pm;
end if;
if exists (select '*' from sysprocedure where proc_name like 'slave_insert_mm') then  
	drop procedure slave_insert_mm;
end if;

	create PROCEDURE "slave_insert_stime"(in table_name char(50), in field_claus char(256) default null, in values_claus char(1000) default null , in select_claus char(1000) default null)
	at 'stime;;ADMIN;slave_insert';

	create PROCEDURE "slave_insert_pm"(in table_name char(50), in field_claus char(256) default null, in values_claus char(1000) default null , in select_claus char(1000) default null)
	at 'pm;;ADMIN;slave_insert';

	create PROCEDURE "slave_insert_mm"(
		in table_name char(50)
		, in field_claus char(256) default null
		, in values_claus char(1000) default null 
		, in select_claus char(1000) default null
		)
	at 'mm;;ADMIN;slave_insert';




if exists (select '*' from sysprocedure where proc_name like 'slave_count_insert_stime') then  
	drop procedure slave_count_insert_stime;
end if;
if exists (select '*' from sysprocedure where proc_name like 'slave_count_insert_pm') then  
	drop procedure slave_count_insert_pm;
end if;
if exists (select '*' from sysprocedure where proc_name like 'slave_count_insert_mm') then  
	drop procedure slave_count_insert_mm;
end if;

	create PROCEDURE slave_count_insert_stime(
			  out inserted integer
			, in table_name char(50)
			, in field_claus char(256) default null
			, in values_claus char(2000) default null 
			, in select_claus char(1000) default null
	)
	at 'stime;;ADMIN;slave_count_insert';

	create PROCEDURE slave_count_insert_pm(
			  out inserted integer
			, in table_name char(50)
			, in field_claus char(256) default null
			, in values_claus char(2000) default null 
			, in select_claus char(1000) default null
	)
	at 'pm;;ADMIN;slave_count_insert';

	create PROCEDURE slave_count_insert_mm(
			  out inserted integer
			, in table_name char(50)
			, in field_claus char(256) default null
			, in values_claus char(2000) default null 
			, in select_claus char(1000) default null
	)
	at 'mm;;ADMIN;slave_count_insert';



--****************************************************************
--                               UPDATE
--****************************************************************
if exists (select '*' from sysprocedure where proc_name like 'slave_update_stime') then  
	drop procedure slave_update_stime;
end if;
if exists (select '*' from sysprocedure where proc_name like 'slave_update_pm') then  
	drop procedure slave_update_pm;
end if;
if exists (select '*' from sysprocedure where proc_name like 'slave_update_mm') then  
	drop procedure slave_update_mm;
end if;


	create PROCEDURE "slave_update_stime"(in table_name char(50), in field_claus char(256) , in values_claus char(1000) , in where_claus char(1000))
	at 'stime;;ADMIN;slave_update';

	create PROCEDURE "slave_update_pm"(in table_name char(50), in field_claus char(256) , in values_claus char(1000) , in where_claus char(1000))
	at 'pm;;ADMIN;slave_update';

	create PROCEDURE "slave_update_mm"(
		in table_name char(50)
		, in field_claus char(256) 
		, in values_claus char(1000) 
		, in where_claus char(1000)
	)
	at 'mm;;ADMIN;slave_update';






if exists (select '*' from sysprocedure where proc_name like 'slave_count_update_stime') then  
	drop procedure slave_count_update_stime;
end if;
if exists (select '*' from sysprocedure where proc_name like 'slave_count_update_pm') then  
	drop procedure slave_count_update_pm;
end if;
if exists (select '*' from sysprocedure where proc_name like 'slave_count_update_mm') then  
	drop procedure slave_count_update_mm;
end if;


	create PROCEDURE slave_count_update_stime(
			  out updated integer
			, in table_name char(50)
			, in field_claus char(256) 
			, in values_claus char(1000) 
			, in where_claus char(1000)
	)
	at 'stime;;ADMIN;slave_count_update';

	create PROCEDURE slave_count_update_pm(
			  out updated integer
			, in table_name char(50)
			, in field_claus char(256) 
			, in values_claus char(1000) 
			, in where_claus char(1000)
	)
	at 'pm;;ADMIN;slave_count_update';

	create PROCEDURE slave_count_update_mm(
			  out updated integer
			, in table_name char(50)
			, in field_claus char(256) 
			, in values_claus char(1000) 
			, in where_claus char(1000)
	)
	at 'mm;;ADMIN;slave_count_update';



--****************************************************************
--                      CURRENCY PROCEDUREIS
--****************************************************************

if exists (select '*' from sysprocedure where proc_name like 'currency_rate_stime') then  
	drop procedure currency_rate_stime;
end if;
if exists (select '*' from sysprocedure where proc_name like 'currency_rate_pm') then  
	drop procedure currency_rate_pm;
end if;

	create PROCEDURE currency_rate_stime(
			out o_date char(20)
			,out o_rate real
			,in p_date char(20) default null
			,in p_id_cur integer default null
	)
	at 'stime;;;slave_currency_rate';

	create PROCEDURE currency_rate_pm(
			out o_date char(20)
			,out o_rate real
			,in p_date char(20) default null
			,in p_id_cur integer default null
	)
	at 'pm;;;slave_currency_rate';
*/


/*

*/
