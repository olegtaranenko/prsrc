update ${orders.table.name} set firmid = ${main.firm.id} where firmid = ${deleted.firm.id};
delete from ${firms.table.name} where firmid = ${deleted.firm.id};
commit;
