call build_table_one_server('jmat', 'markmaster', 1);
call build_table_one_server('jmat', 'accountn', 1);

begin
	declare nc int;
	declare ic int;
	declare ic0 int;
	declare ex int; declare c1 int; declare c0 int;declare cc1 int; declare cc0 int;
	set nc = 0; set ic = 0; set ic0 = 0; set c1 = 0; set c0 = 0;set cc1 = 0; set cc0 = 0;
	for c_rmt_servers as a dynamic scroll cursor for
		select c.id as r_id, char_length(nu) as r_lnu, nu as r_nu
		from jmat_markmaster c 
--		join sdocs n on n.id_jmat = c.id
--		where c.id in (57378,57365,57380,57381,54890, 54919, 57341)
--		where c.dat > '20051013'
	do
		select count(*) into ex from jmat_stime where id = r_id;
		if ex = 1 then
			set nc = nc + 1;
		else
			select count(*) into ex from jmat_accountn where id = r_id;
			if ex = 1 then
				set ic = ic + 1;
			else
				set ic0 = ic0 + 1;
				if r_lnu = 7 and charIndex('-', r_nu) = 0 then
					select count(*) into ex from sdocs where r_id = id_jmat;
					if ex = 1 then
						set cc1 = cc1 + 1;
					else
--message r_id, ' ', r_nu to client;
						set cc0 = cc0 + 1;
						delete from jmat_markmaster where id = r_id;
					end if;
					set c1 = c1 + 1;
				else
					set c0 = c0 + 1;
				end if;
			end if
		end if;
	end for;
--	message nc, ' ', ic, ' ', ic0, ' ', c1, ' ', c0, ' ', cc1, ' ', cc0 to client;
end;


begin
	declare nc int;
	declare ic int;
	declare ic0 int;
	declare ex int; declare c1 int; declare c0 int;declare cc1 int; declare cc0 int;
	set nc = 0; set ic = 0; set ic0 = 0; set c1 = 0; set c0 = 0;set cc1 = 0; set cc0 = 0;
	for c_rmt_servers as a dynamic scroll cursor for
		select c.id as r_id, char_length(nu) as r_lnu, nu as r_nu
		from jmat_accountn c 
--		join sdocs n on n.id_jmat = c.id
--		where c.id in (57378,57365,57380,57381,54890, 54919, 57341)
		where c.dat > '20051013'
	do
		select count(*) into ex from jmat_stime where id = r_id;
		if ex = 1 then
			set nc = nc + 1;
		else
			select count(*) into ex from jmat_markmaster where id = r_id;
			if ex = 1 then
				set ic = ic + 1;
			else
				set ic0 = ic0 + 1;
				if r_lnu = 7 and charIndex('-', r_nu) = 0 then
					select count(*) into ex from sdocs where r_id = id_jmat;
					if ex = 1 then
						set cc1 = cc1 + 1;
					else
--message r_id, ' ', r_nu to client;
						set cc0 = cc0 + 1;
						delete from jmat_accountn where id = r_id;
					end if;
					set c1 = c1 + 1;
				else
					set c0 = c0 + 1;
				end if;
			end if
		end if;
	end for;
--message nc, ' ', ic, ' ', ic0, ' ', c1, ' ', c0, ' ', cc1, ' ', cc0 to client;
end;


call build_table_one_server('jmat', 'markmaster', 0);
call build_table_one_server('jmat', 'accountn', 0);

