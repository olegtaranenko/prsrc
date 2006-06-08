if exists (select '*' from sysprocedure where proc_name like 'change_mat_qty') then
	drop procedure change_mat_qty;
end if;
	-- процедура вызывается для коррекции суммы по позиции
	-- при изменении количества в интерфейсе stime

create procedure change_mat_qty (
	  in p_id_mat integer
	, in p_new_quant float
)
begin
	declare v_kol1 float;
	select kol1 into v_kol1 from mat where id= p_id_mat;
	if abs(round(v_kol1, 3)) > 0.001 then
		update mat set 
			 summa       = summa       / kol1 * p_new_quant
			,summa_sale  = summa_sale  / kol1 * p_new_quant
			,summav      = summav      / kol1 * p_new_quant
			,summa_salev = summa_salev / kol1 * p_new_quant
			,kol1        = p_new_quant
			,kol2        = p_new_quant
			,kol3        = p_new_quant
		where id = p_id_mat;
	else
		update mat set 
			kol1        = p_new_quant
			,kol2        = p_new_quant
			,kol3        = p_new_quant
		where id = p_id_mat;
	end if;
end;



if exists (select '*' from sysprocedure where proc_name like 'order_import') then
	drop procedure order_import;
end if;

create procedure order_import (
-- процедура должна вызываться при смене типа накладной с рублевой
-- на импортную или наоборот
-- Пересчет денежных значений по позициям должен осуществляться 
-- вне этой функции.
	  in p_id_jmat integer
	, in p_currency_id integer
	, in p_id_guide integer
	, in p_tp1 integer
	, in p_tp2 integer
	, in p_tp3 integer
	, in p_tp4 integer
) 
begin

	declare out_cur_date varchar(20);
	declare v_rate float;


	if p_currency_id is not null then
		-- текущий курс валюты
		call slave_currency_rate(out_cur_date, v_rate, null, p_currency_id);

		-- 
		update jmat set 
			id_guide = p_id_guide
			, id_curr = p_currency_id 
			, tp1 = p_tp1
			, tp2 = p_tp2
			, tp3 = p_tp3
			, tp4 = p_tp4
			, curr = isnull(v_rate, 1.0)
		where 
			id = p_id_jmat;

/*
		update mat set
			summa_salev = kol1 * v_rate
		where id_jmat = p_id_jmat;
*/
		end if;


end;

