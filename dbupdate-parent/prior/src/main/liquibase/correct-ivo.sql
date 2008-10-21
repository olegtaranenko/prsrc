begin
	raiserror 17000 'test raiserror!';
	truncate table sdmcventure;
	truncate table sdocsventure;

	call bootstrap_blocking();
--	call cre_block_var('supress_cum_update');
--	call cre_block_var('supress_diary_update');

--	call cre_block_var('blocks_inited');

	call ivo_generate(10);
end;
