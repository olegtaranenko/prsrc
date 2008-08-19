alter table compl drop foreign key compl_2_inv;
alter table compl drop foreign key compl_2_inv_belong;

alter table compl add foreign key compl_2_inv (id_inv) references inv(id) on update cascade on delete cascade;
alter table compl add foreign key compl_2_inv_belong (id_inv_belong) references inv(id) on update cascade on delete cascade;
