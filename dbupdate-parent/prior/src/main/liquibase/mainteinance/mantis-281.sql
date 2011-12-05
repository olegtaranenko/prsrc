update orders set statusid = 6 where numorder in (11082513, 11101003);
delete from sdocs where numdoc = 9291008;
commit;