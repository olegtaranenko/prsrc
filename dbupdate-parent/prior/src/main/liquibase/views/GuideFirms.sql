if exists (select 1 from sysviews where viewname = 'GuideFirms') then
	drop view GuideFirms
end if;
 

CREATE VIEW
	GuideFirms
AS
SELECT
    FirmId,
    Name,
    xLogin,
    FIO,
    Phone,
    Kategor,
    Sale,
    ManagId,
    Address,
    Fax,
    Email,
    Atr1,
    Atr2,
    Atr3,
    Pass,
    Level,
    Type,
    Katalog,
    year01,
    year02,
    year03,
    year04,
    id_voc_names
FROM
    FirmGuide
where 
	WerkId = 2;

