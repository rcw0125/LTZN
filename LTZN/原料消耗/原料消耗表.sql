DROP TABLE 原料消耗;
CREATE TABLE 原料消耗
(
      日期   　　　     DATE NOT NULL,
      高炉    　　　　  NUMBER(2) NOT NULL,
      机烧矿　　　　　  NUMBER(20,4),
      竖球　　　　　　　NUMBER(20,4),
      本溪矿　　　　　　NUMBER(20,4),
      工艺称　　　　　　NUMBER(20,4),
      焦丁　　　　　　　NUMBER(20,4),
      煤粉　　　　　　　NUMBER(20,4),
      富氧量　　　　　　NUMBER(20,4),
      自产湿焦　　　　　NUMBER(20,4),
      落地湿焦　　　　　NUMBER(20,4),
      熟料　　          NUMBER(20,4),
　　　熟料名称　　　　　VARCHAR2(20),
      生料        　　  NUMBER(20,4),
      生料名称　　　　　VARCHAR2(20),
      CONSTRAINT "原料消耗主键" PRIMARY KEY(日期,高炉)
);