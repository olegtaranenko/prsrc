if exists (select 1 from systriggers where trigname = 'compl_save_identity' and tname = 'compl') then 
	drop trigger compl.compl_save_identity;
end if;


CREATE TRIGGER "compl_save_identity".compl_save_identity after insert order 1 on
DBA.compl
for each row
begin
call set_last_identity('compl')
end
