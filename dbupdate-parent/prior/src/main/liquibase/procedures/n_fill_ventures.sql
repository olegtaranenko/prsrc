ALTER PROCEDURE "DBA"."n_fill_ventures" (
	  p_filterId     integer
	, p_begin         date  default null
	, p_end           date  default null
)
begin
	insert into #periods (ventureId, label)
	select ventureId, ventureName 
	from guideVenture;
end