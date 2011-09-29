ALTER PROCEDURE "DBA"."n_fill_klasses" (
	  p_filterId     integer
	, p_begin         date  default null
	, p_end           date  default null
)
begin

	declare v_table_name    varchar(64);
	declare v_ord_table varchar(64);

	declare v_sql long varchar;

	create table #sale_item (
		 numorder    integer
		,nomnom      varchar(20)
		,materialQty         float
		,sm          float
		,inDate      date
		,firmId      integer
		,klassid     integer
	);

	set v_table_name = 'sGuideKlass';
	set v_ord_table = get_tmp_ord_table_name(v_table_name);
	execute immediate get_tmp_ord_create_sql(v_ord_table); -- #sGuideKlass_ord


	call n_internal_klasses (p_begin, p_end, v_table_name, null, null);

end
