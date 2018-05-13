DROP Table 慢风;
Create Table  慢风
(
   时间       DATE NOT NULL,
   高炉       NUMBER(2) NOT NULL,
   间隔       NUMBER(20,4),
   分类       VARCHAR2(100),
   原因       VARCHAR2(100),
   CONSTRAINT "慢风主键" PRIMARY KEY(时间,高炉) 
);