alter table system add (total_accounting_date timestamp, venture_anl_id integer);

update system set total_accounting_date = null, venture_anl_id = 3;

update guideventure set activity_start = null where ventureid = 3;

commit;
