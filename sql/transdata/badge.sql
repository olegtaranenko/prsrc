create table sBadge (
	productid smallint
	, nomnom varchar(20)
	, id_size integer
--	, classId integer
--	, nm varchar(32)
	, id_compl integer
	, primary key (productid, nomnom, id_size)
);


alter table sBadge add constraint b_product foreign key (productid, nomnom) references sproducts (productid, nomnom) on update cascade

alter table size modify id_size not null;

alter table size add primary key (id_size);

alter table sBadge add constraint fk_sz foreign key (id_size) references size (id_size) on update cascade;

alter table sProducts add isbadge tinyint;

update sproducts set isbadge = 0;

alter table sproducts modify isbadge not null default 0;

