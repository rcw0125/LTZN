DROP Table 日原料成分;
Create Table 日原料成分
(
    P01日期       DATE NOT NULL,
    P02名称       VARCHAR2(20) NOT NULL,
    P03TFE        NUMBER(20,4),
    P04SIO2       NUMBER(20,4),
    P05CAO        NUMBER(20,4),
    P06FEO        NUMBER(20,4),
    P07MGO        NUMBER(20,4),
    P08S          NUMBER(20,4),
    P09R          NUMBER(20,4),
    CONSTRAINT "日原料成分主键" PRIMARY KEY(P01日期,P02名称)
);