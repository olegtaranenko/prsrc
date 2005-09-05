if exists (select 1 from sysforeignkeys where foreign_creator = 'dba' 
and foreign_tname = ''
and primary_tname = ''
) then
end if;



alter table sproducts drop foreign key sproducts_856;
alter table sguideproducts drop foreign key sguideproducts_857;
alter table sguideproducts drop foreign key sGuideFormulssGuideProducts;
alter table sguidenomenk drop foreign key sguideformulssguidenomenk;
alter table sguidenomenk drop foreign key sguideformulssguidenomenk1;
alter table sdocs drop foreign key sguidesourcesdocs;
alter table sdocs drop foreign key sguidesourcesdocs1;
alter table xvariantnomenc drop foreign key sProductsxVariantNomenc;
alter table sguideklass drop foreign key sguideklass_851;
alter table sguideseries drop foreign key sguideseries_858;
alter table ybook drop foreign key yBook_381;
alter table ybook drop foreign key yBook_382;
alter table ybook drop foreign key yBook_384;
alter table ybook drop foreign key yBook_385;
alter table yGuideDetail drop foreign key yGuideDetyGuideDetail;
alter table yGuideDetail drop foreign key yGuideDetail_383;
alter table yGuidePurpose drop foreign key yGuidePurpose_386;
alter table yGuidePurpose drop foreign key yGuidePurpose_387;
alter table yGuidePurpose drop foreign key yGuidePurpyGuidePurpose;





alter table sproducts add constraint sproducts_856 foreign key (productid) references sguideproducts(prid) on update cascade;
alter table sguideproducts add constraint sguideproducts_857 foreign key (prSeriaId) references sguideseries(seriaid) on update cascade;
alter table sguideproducts add constraint sGuideFormulssGuideProducts foreign key (formulaNom) references sGuideFormuls(nomer) on update cascade;
alter table sguidenomenk add constraint sguideformulssguidenomenk foreign key (formulaNom) references sGuideFormuls(nomer) on update cascade;
alter table sguidenomenk add constraint sguideformulssguidenomenk1 foreign key (formulaNomW) references sGuideFormuls(nomer) on update cascade;
alter table sdocs add constraint sguidesourcesdocs foreign key (SourId) references sguidesource(sourceId) on update cascade;
alter table sdocs add constraint sguidesourcesdocs1 foreign key (DestId) references sguidesource(sourceId) on update cascade;
alter table sguideklass add constraint sguideklass_851 foreign key (parentklassId) references sguideklass(klassId) on update cascade;
alter table sguideseries add constraint sguideseries_858 foreign key (parentseriaId) references sguideseries(seriaId) on update cascade;

alter table sproducts delete prid;
alter table sguideproducts delete seriaid;
alter table sguideproducts delete nomer;
alter table sguidenomenk delete nomer;
alter table sdocs delete sourceid;
alter table xvariantnomenc delete productid;
alter table ybook delete pId;
alter table ybook delete id;
alter table ybook delete number;
alter table ybook delete subNumber;
alter table yGuideDetail delete pId;
alter table yGuidePurpose delete number;
alter table yGuidePurpose delete subNumber;
alter table yGuidePurpose delete descript;
