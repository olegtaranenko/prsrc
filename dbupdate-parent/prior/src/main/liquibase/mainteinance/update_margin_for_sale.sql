begin
declare signIndex int;
declare sign char(1);
declare cNumber varchar(20);	
declare num real;
declare v_margin real;
declare r_nomer int;
declare r_formula varchar(100);
for x as cx dynamic scroll cursor for
	select 
		n.nomnom as r_nomnom, n.formulaNomW as r_formulaNomer
	from sguidenomenk n
	for update
do
	select f.nomer , f.formula 
	into r_nomer, r_formula
	from sguideformuls f
	where f.nomer = r_formulaNomer;

	set v_margin = 0;
	if isnull(char_length(r_formula), 0) != 0 then
		set signIndex = charindex('/', r_formula);
		if signIndex = 0 then
			set signIndex = charindex('*', r_formula);
			set sign = '*';
		else 
			set sign = '/';
		end if;
		if signIndex > 0 then
			set cNumber = substring (r_formula, signIndex + 1);
			set num = convert(real, cNumber);
		else
			set cNumber = '?';
			set num = 1;
		end if;
		if sign = '/' then
			set v_margin = (1 - num) * 100;
		elseif sign = '*' then
			set v_margin = ( 1 - 1 / num) * 100;
		end if;
	end if;
//	message r_formula, ' - ', cNumber, ', ', r_note to client;
//	message sign, ', ', signIndex, ', ', v_margin to client;
	update sGuideNomenk set margin = round(v_margin, 1) where current of cx;
	commit;
end for;
