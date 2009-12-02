if exists (select 1 from systriggers where trigname = 'wf_delete_sdocs' and tname = 'sdocs') then 
	drop trigger sdocs.wf_delete_sdocs;
end if;

create TRIGGER wf_delete_sdocs before delete on
sdocs
referencing old as old_name
for each row
begin
	declare remoteServer varchar(32);
	declare no_echo integer;

	set no_echo = 0;

  	begin
  		message '@stime_sdocs = ', @stime_sdocs to log;
		select @stime_sdocs into no_echo; 
	exception 
		when other then
			set no_echo = 0;
	end;

	if no_echo = 1 then
		return;
	end if;



	if (old_name.id_jmat is not null) then
		call block_remote('stime', get_server_name(), 'jmat');
		call block_remote('stime', get_server_name(), 'mat');
		call delete_remote('stime', 'jmat', 'id = ' + convert(varchar(20), old_name.id_jmat));
		call unblock_remote('stime', get_server_name(), 'jmat');
		call unblock_remote('stime', get_server_name(), 'mat');
	end if;

	select sysname into remoteServer 
	from  guideventure v 
	where old_name.ventureId = v.ventureId and v.standalone = 0;

--	message 'remoteServer = ', remoteServer to client;
	if remoteServer is not null and remoteServer != 'stime' then
		call block_remote(remoteServer, get_server_name(), 'jmat');
		call block_remote(remoteServer, get_server_name(), 'mat');
		call delete_remote(remoteServer, 'jmat', 'id = ' + convert(varchar(20), old_name.id_jmat));
		call unblock_remote(remoteServer, get_server_name(), 'jmat');
		call unblock_remote(remoteServer, get_server_name(), 'mat');
	end if;
end;


