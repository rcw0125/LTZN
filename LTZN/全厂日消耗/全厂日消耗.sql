DROP Table 全厂日消耗;
Create Table 全厂日消耗
(
    P01日期             DATE NOT NULL,
    P02自产焦水分       NUMBER(20,4),
    P03落地焦水分       NUMBER(20,4),
    P04焦粉水分         NUMBER(20,4),
    P05二号皮带         NUMBER(20,4),
    P06三号皮带         NUMBER(20,4),
    P07总返矿           NUMBER(20,4),
    P08烧结二号称       NUMBER(20,4),
    P09备注1            VARCHAR2(100),
    P10备注2            VARCHAR2(100),
    CONSTRAINT "全厂日消耗主键" PRIMARY KEY(P01日期)
);