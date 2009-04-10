if exists (select 1 from systriggers where trigname = 'newid' and tname = 'compl') then 
	drop trigger compl.newid;
end if;

create TRIGGER "newid" before insert order 2 on
DBA.compl
referencing new as new_name
for each row begin
	declare ll_id integer;
	set ll_id=new_name.id;
	if ll_id is null then
		set new_name.Id=get_next_id('compl','Id');
		set @@id=new_name.id;
		call set_last_identity2('compl')
	end if
end