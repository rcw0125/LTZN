drop table manfeng;
create table manfeng
(
sj              date not null,
gaolu           number(2) not null,
jiange          number(10),
fenlei          varchar2(100),
yuanyin         varchar2(100),
CONSTRAINT "manfengpk" PRIMARY KEY(sj,gaolu)
)