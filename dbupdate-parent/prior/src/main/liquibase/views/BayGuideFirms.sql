if exists (select 1 from sysviews where viewname = 'BayGuideFirms') then
	drop view BayGuideFirms
end if;
 

CREATE VIEW
	BayGuideFirms
AS
SELECT
    FirmId
    ,Name
    ,xLogin
    ,FIO
    ,Phone
    ,Sale
    ,ManagId
    ,Fax
    ,Email
    ,Pass
    ,Type
    ,Katalog
    ,year01
    ,year02
    ,year03
    ,year04
    ,Kontakt
    ,Otklik
    ,id_voc_names
    ,regionid
    ,oborudId
    ,bayStatusId
    ,tools
    ,Address
    ,City
    ,Delat
    ,InceptionYear
    ,Director
    ,KvoSotr
    ,Steuer
    ,HomePage
    ,Supplier
FROM
     FirmGuide
where 
	WerkId = 1
