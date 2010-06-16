if exists (select 1 from systriggers where trigname = 'wf_delete_orders' and tname = 'Orders') then 
	drop trigger Orders.wf_delete_orders;
end if;

create TRIGGER wf_delete_orders before delete on
Orders
referencing old as old_name
for each row
begin
	declare remoteServer varchar(32);
	declare deleted integer;
	select sysname into remoteServer from guideventure where ventureId = old_name.ventureId;
	if remoteServer is not null then
		-- в комтехе констрейнт не каскадный - жалко
		-- удаляем договор явно
		call purge_jscet(remoteServer, old_name.id_jscet);
	end if;
--  delete from inv where id = old_name.id_inv;
end;


