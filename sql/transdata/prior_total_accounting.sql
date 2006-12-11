
-- в продакш => нет
-- Пересчитать вступительную инвентаризацию в аналитике
call bootstrap_blocking();
--call inventory_order('20051013 21:00', 1, null);

-- перерасчитать все цены на текущие
--call wf_cost_bulk_change(0);

-- в продакш => нет
-- инвентаризация по предприятиям (пм, мм) на текущую дату
--call v_inventory_order(null, '20041030 23:00');
--call v_inventory_order();

--Исправление взаимозачетов, (с учетом Аналитики)
truncate table sdmcventure;
truncate table sdocsventure;

call ivo_generate(10);

commit;
exit;

begin
	for crs as c dynamic scroll cursor for
		select id as r_ivo_id from sdocsventure where cumulative_id is null
	do
		call ivo_to_comtex(r_ivo_id);
	end for;
end;
