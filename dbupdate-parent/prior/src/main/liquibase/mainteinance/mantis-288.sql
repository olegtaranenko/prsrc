

update orders 
set statusId = ordersequip.statusEquipId
from ordersequip
where orders.numorder = ordersequip.numorder
and orders.statusId != ordersEquip.statusEquipId
and orders.statusId != 6


update ordersequip set stat = '�����' where statusEquipId = 4 and Stat != '�����'
commit
