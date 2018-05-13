DROP TABLE 炉铁;
CREATE TABLE 炉铁
{   高炉                NUMBER(2)  not null,
    正点时间            DATE not null,            
    对罐时间            DATE,
    开口时间            DATE,
    堵口时间            DATE,
    晚点时间            DATE,
    通知时间            DATE,
    实际产量            NUMBER(10, 2)          
    理论产量            NUMBER(10, 2)
    折算产量            NUMBER(10, 2)
    估计产量            NUMBER(10, 2)
    估计罐数            NUMBER(10, 2)
    铁水温度
    深度
    角度
    打泥量
    料批数
    去处
    FEC                 NUMBER(10, 2),
    FESI                NUMBER(10, 2), 
    FEMN                NUMBER(10, 2), 
    FES                 NUMBER(10,  2),
    FEP                 NUMBER(10, 2), 
    FETI                NUMBER(10, 2), 
    估计FESI            NUMBER(10, 2), 
    估计FES             NUMBER(10, 2), 
    渣样                VARCHAR2(20), 
    水温1               NUMBER(10, 2), 
    水温2               NUMBER(10, 2), 
    ZHASIO2             NUMBER(10, 2), 
    ZHACAO              NUMBER(10, 2), 
    ZHAMGO              NUMBER(10, 2), 
    ZHAAL2O3            NUMBER(10, 2), 
    ZHAS                NUMBER(10, 2),
    ZHATIO2                NUMBER(10, 2),
    ZHAR2               NUMBER(10, 2), 
    总炉次  
    班次
    班炉次
    分类1
    分类2
    分类3
 }
    