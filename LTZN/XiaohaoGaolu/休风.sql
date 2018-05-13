DROP Table 休风;
Create Table  休风
(
   时间       DATE NOT NULL,
   高炉       NUMBER(2) NOT NULL,
   间隔       NUMBER(20,4),
   分类       VARCHAR2(100),
   原因       VARCHAR2(100),
   CONSTRAINT "休风主键" PRIMARY KEY(时间,高炉) 
);