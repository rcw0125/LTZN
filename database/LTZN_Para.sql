conn lf/lf@liantie;
set echo on;
spool log/LTZN_Para.log;
drop table LTZN_Para;
create table LTZN_Para(
 Name varchar2(1000),
 Val number(18,8),
 StrVal varchar2(1000),
 Constraint LTZN_Para_Primary_Key Primary Key(Name)
 );
insert into LTZN_Para(Name,Val) values('Si最大',0.7);
insert into LTZN_Para(Name,Val) values('Si次大',0.5);
insert into LTZN_Para(Name,Val) values('Si次小',0.3);
insert into LTZN_Para(Name,Val) values('Si最小',0.2);
insert into LTZN_Para(Name,Val) values('S单值上限',0.030021);
insert into LTZN_Para(Name,Val) values('S单值均值',0.02024);
insert into LTZN_Para(Name,Val) values('S单值下限',0.010027);
insert into LTZN_Para(Name,Val) values('S极差上限',0.01225);
insert into LTZN_Para(Name,Val) values('S极差均值',0.00375);
insert into LTZN_Para(Name,Val) values('S极差下限',0);
insert into LTZN_Para(Name,Val) values('R最大',1.25);
insert into LTZN_Para(Name,Val) values('R次大',1.20);
insert into LTZN_Para(Name,Val) values('R次小',1.07);
insert into LTZN_Para(Name,Val) values('R最小',1);
insert into LTZN_Para(Name,Val) values('Si显示最大值',1);
insert into LTZN_Para(Name,Val) values('Si显示最小值',0);
insert into LTZN_Para(Name,Val) values('S显示最大值',1);
insert into LTZN_Para(Name,Val) values('S显示最小值',0);
insert into LTZN_Para(Name,Val) values('R显示最大值',1);
insert into LTZN_Para(Name,Val) values('R显示最小值',0);
spool off;
exit;