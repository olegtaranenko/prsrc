begin
	declare v_count int; declare v_wrong_outcome int;
	declare v_data_lock int;
	declare v_start_total datetime;

	select total_accounting_date into v_start_total from system;

	set v_wrong_outcome = 0;
	set v_count = 0;
mainloop:	
	for c_bay as b dynamic scroll cursor for
		select 
			numorder as r_numorder
			, id_jscet as r_id_jscet
			, invoice as r_invoice
			, inDate as r_inDate
			, v.sysname as r_server
		from orders o
		join guideventure v on v.ventureId = o.ventureid
		where inDate >= v_start_total 
			union
		select 
			numorder as r_numorder
			, id_jscet as r_id_jscet
			, invoice as r_invoice
			, inDate as r_inDate
			, v.sysname as r_server
		from bayorders o
		join guideventure v on v.ventureId = o.ventureid
		where inDate >= v_start_total 
		order by 1 desc
	do
		set v_data_lock = select_remote(r_server, 'jscet', 'data_lock', 'id = ' + convert(varchar(10), r_id_jscet));
		if v_data_lock = 1 then
--			message 'numorder = ', r_numorder, ': invoice = ', r_invoice, ': id_jscet = ', r_id_jscet, ': v_data_lock = ', v_data_lock to client;
			set v_wrong_outcome = 0;
			sdmcloop:
			for c_sdmc as csdmc dynamic scroll cursor for
				select numext as r_numext, id_jmat as r_id_jmat 
				from sdocs where numdoc = r_numorder
			do
				set v_wrong_outcome = select_remote(r_server, 'jmat', 'count(*)', 'id_jscet = ' + convert(varchar(10), r_id_jscet) + ' and tp1=3 and tp2=2 and tp3=1');
				if v_wrong_outcome > 0 then
					message 'numorder = ', r_numorder, ': invoice = ', r_invoice, ': id_jscet = ', r_id_jscet, ': r_indate = ', r_indate to client;
					message '	numext = ', r_numext, ': id_jmat = ', r_id_jmat to client;
					set v_count = v_count + 1;
					call update_remote(r_server, 'jmat', 'tp1', '2', 'id_jscet = ' + convert(varchar(10), r_id_jscet));
--					call update_remote(r_server, 'jmat', 'tp2', '2', 'id_jscet = ' + convert(varchar(10), r_id_jscet));
					call update_remote(r_server, 'jmat', 'tp3', '2', 'id_jscet = ' + convert(varchar(10), r_id_jscet));
					call update_remote(r_server, 'jmat', 'tp4', '0', 'id_jscet = ' + convert(varchar(10), r_id_jscet));
					leave sdmcloop;
				end if;
			end for;
		end if;
		if v_count > 2000 then
			--leave mainloop;
		end if;
	end for;
	message 'total rows: ', v_count to client;
end;

commit;


/*
numorder = 9101606: invoice = 501566: r_server = accountN: r_indate = 2009-10-16 13:44:33.000
numorder = 9073007: invoice = 55634: r_server = markmaster: r_indate = 2009-07-30 15:13:23.000
numorder = 9072815: invoice = 501062: r_server = accountN: r_indate = 2009-07-28 12:41:06.000
numorder = 9062316: invoice = 50876: r_server = accountN: r_indate = 2009-06-23 17:04:00.000
numorder = 9060913: invoice = 55496: r_server = markmaster: r_indate = 2009-06-09 14:40:37.000
numorder = 9052219: invoice = 50691: r_server = accountN: r_indate = 2009-05-22 15:51:16.000
numorder = 9051813: invoice = 50654: r_server = accountN: r_indate = 2009-05-18 15:21:18.000
numorder = 9051209: invoice = 50619: r_server = accountN: r_indate = 2009-05-12 13:13:46.000
numorder = 9042127: invoice = 50531: r_server = accountN: r_indate = 2009-04-21 17:49:45.000
numorder = 9031217: invoice = 55204: r_server = markmaster: r_indate = 2009-03-12 17:06:27.000
numorder = 8102814: invoice = 50904: r_server = accountN: r_indate = 2008-10-28 12:57:09.000
numorder = 8102007: invoice = 552394: r_server = markmaster: r_indate = 2008-10-20 12:17:19.000
numorder = 8101011: invoice = 50852: r_server = accountN: r_indate = 2008-10-10 12:46:09.000
numorder = 8091832: invoice = 552115: r_server = markmaster: r_indate = 2008-09-18 16:47:51.000
numorder = 8090412: invoice = 50703: r_server = accountN: r_indate = 2008-09-04 14:15:40.000
numorder = 8090302: invoice = 50697: r_server = accountN: r_indate = 2008-09-03 11:37:59.000
numorder = 8082801: invoice = 551919: r_server = markmaster: r_indate = 2008-08-28 10:25:19.000
numorder = 8082111: invoice = 551857: r_server = markmaster: r_indate = 2008-08-21 13:01:32.000
numorder = 8080515: invoice = 551737: r_server = markmaster: r_indate = 2008-08-05 14:15:52.000
numorder = 8072107: invoice = 551613: r_server = markmaster: r_indate = 2008-07-21 13:57:48.000
numorder = 8030506: invoice = 55500: r_server = markmaster: r_indate = 2008-03-05 12:26:46.000
numorder = 8020420: invoice = 55241: r_server = markmaster: r_indate = 2008-02-04 15:52:35.000
numorder = 7121030: invoice = 501387: r_server = accountN: r_indate = 2007-12-10 14:05:45.000
numorder = 7081319: invoice = 551624: r_server = markmaster: r_indate = 2007-08-13 15:11:15.000
numorder = 7080716: invoice = 50835: r_server = accountN: r_indate = 2007-08-07 16:25:49.000
numorder = 7060804: invoice = 551179: r_server = markmaster: r_indate = 2007-06-08 10:44:43.000
*/
