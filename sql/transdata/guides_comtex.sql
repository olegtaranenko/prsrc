create table DBA.guide_806 (id integer, nm varchar(99))

insert into dba.guide_806 (id, nm) select 0, ''


insert into guides (id, kind, nm, namer)
select 806, 'G__W_GUIDE', 'GUIDE_806', 'Интегарция с Pior'


insert into browsers (
Id_GUIDES,NU,NM,NAMER,mask,SUM,EN_SCREEN,EN_PRINT,FIELD_TYPE,FIELD_ATTR,EDITABLE,VER,added,sel_str,grp,parent_col_name,child_col_name,edit_id_guide,subtuner_retrieve_args,ini_value_add_row,standart_nm,supress,group_by,header_id,col_num,unique_col,unique_move,group_id,cursor_after_edit,valid_expression,auto_ini_val,on_update_func
)
select 806,1,'GUIDE_806.ID','Идентификатор','','','0','0','','','0','','0','','','','',0,'','','ID','0','0',0,'','0','0',0,'','',null,''
union 
select 806,2,'GUIDE_806.NM','Наименование','','','1','1','','','1','','1','','','','',0,'','','NM','0','0',0,'','0','0',0,'','',null,''

select * from browsers where id_guides in (805, 806)

create TRIGGER "GUIDE_806_save_identity".GUIDE_806_save_identity 
after insert order 1 on DBA.GUIDE_806
for each row 
begin call set_last_identity('GUIDE_806') end;
