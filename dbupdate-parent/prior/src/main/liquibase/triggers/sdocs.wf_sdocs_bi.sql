if exists (select 1 from systriggers where trigname = 'wf_sdocs_bi' and tname = 'sdocs') then 
	drop trigger sdocs.wf_sdocs_bi;
end if;

create 
	trigger wf_sdocs_bi before insert on 
sdocs
referencing new as new_name
for each row
begin
	declare v_id_jmat integer;
	declare v_venture_id integer;


	call wf_dual_distribute (
		new_name.numdoc
		, new_name.numext
		, new_name.sourid
		, new_name.destid
		, new_name.xdate
		, v_id_jmat
		, v_venture_id 
	);
	set new_name.id_jmat = v_id_jmat;
	set new_name.ventureId = v_venture_id;

end;

