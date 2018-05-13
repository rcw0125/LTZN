namespace LTZN.技术日报 {
    using System.ComponentModel;

    partial class 技术日报数据集
    {
        partial class 技术日报DataTable
        {
            public 技术日报Row Add技术日报Row(
                            System.DateTime P日期,
                            string P01单位,
                            string P02项目)
            {
                技术日报Row row技术日报Row = ((技术日报Row)(this.NewRow()));
                row技术日报Row.ItemArray = new object[] {
                        P日期,
                        P01单位,
                        P02项目,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull};
                this.Rows.Add(row技术日报Row);
                return row技术日报Row;
            }

            public 技术日报内容 技术日报内容
            {
                get
                {
                    if (this.Count == 12)
                    {
                        技术日报内容 n = new 技术日报内容(this[0].P日期);
                        foreach (技术日报Row r in this)
                        {
                            技术日报内容项 x = new 技术日报内容项(r.P01单位, r.P02项目);
                            x.P03合格铁 = r.PA03合格铁;
                            x.P04炼钢铁 = r.PA04炼钢铁;
                            x.P05铸造铁 = r.PA05铸造铁;
                            x.P06号外铁 = r.PA06号外铁;
                            x.P07合格率 = r.PA07合格率;
                            x.P08高炉利用系数 = r.PA08高炉利用系数;
                            x.P09高炉实物系数 = r.PA09高炉实物系数;
                            x.P10原料矿合计总耗 = r.PA10原料矿合计总耗;
                            x.P11原料矿合计单耗 = r.PA11原料矿合计单耗;
                            x.P12原料矿机烧 = r.PA12原料矿机烧;
                            x.P13原料矿竖炉球 = r.PA13原料矿竖炉球;
                            x.P14原料矿CT = r.PA14原料矿CT;
                            x.P15原料矿其它熟料 = r.PA15原料矿其它熟料;
                            x.P16原料矿本溪矿 = r.PA16原料矿本溪矿;
                            x.P17原料矿其它生料= r.PA17原料矿其它生料;
                            x.P18废铁总耗 = r.PA18废铁总耗;
                            x.P19废铁单耗 = r.PA19废铁单耗;
                            x.P20回收率 = r.PA20回收率;
                            x.P21熟料比 = r.PA21熟料比;
                            x.P22平均风温 = r.PA22平均风温;
                            x.P23炉顶温度 = r.PA23炉顶温度;
                            x.P24热风压力 = r.PA24热风压力;
                            x.P25炉顶压力 = r.PA25炉顶压力;
                            x.P26富氧率 = r.PA26富氧率;
                            x.P27入炉焦炭总耗 = r.PA27入炉焦炭总耗;
                            x.P28入炉焦炭单耗 = r.PA28入炉焦炭单耗;
                            x.P29煤粉总耗 = r.PA29煤粉总耗;
                            x.P30煤粉单耗 = r.PA30煤粉单耗;
                            x.P31焦丁总耗 = r.PA31焦丁总耗;
                            x.P32焦丁单耗 = r.PA32焦丁单耗;
                            x.P33综合焦炭总耗 = r.PA33综合焦炭总耗;
                            x.P34综合焦炭单耗 = r.PA34综合焦炭单耗;
                            x.P35综合折算焦比 = r.PA35综合折算焦比;
                            x.P36冶炼强度 = r.PA36冶炼强度;
                            x.P37焦炭负荷 = r.PA37焦炭负荷;
                            x.P38干毛焦 = r.PA38干毛焦;
                            x.P39炼钢铁SI = r.PA39炼钢铁SI;
                            x.P40炼钢铁MN = r.PA40炼钢铁MN;
                            x.P41炼钢铁P = r.PA41炼钢铁P;
                            x.P42炼钢铁S = r.PA42炼钢铁S;
                            x.P43铸造铁SI = r.PA43铸造铁SI;
                            x.P44铸造铁MN= r.PA44铸造铁MN;
                            x.P45铸造铁P = r.PA45铸造铁P;
                            x.P46铸造铁S = r.PA46铸造铁S;
                            x.P47炉渣碱度 = r.PA47炉渣碱度;
                            x.P48休风情况 = r.PA48休风情况;
                            x.P49慢风 = r.PA49慢风;
                            x.P50坐料次数 = r.PA50坐料次数;
                            x.P51悬料次数 = r.PA51悬料次数;
                            x.P52崩料次数 = r.PA52崩料次数;
                            x.P53风口损坏数大 = r.PA53风口损坏数大;
                            x.P54风口损坏数中 = r.PA54风口损坏数中;
                            x.P55风口损坏数小 = r.PA55风口损坏数小;
                            x.P56渣口损坏数大 = r.PA56渣口损坏数大;
                            x.P57渣口损坏数中 = r.PA57渣口损坏数中;
                            x.P58渣口损坏数小 = r.PA58渣口损坏数小;
                            x.P59本厂铸造SI大 = r.PA59本厂铸造SI大;
                            x.P60本厂铸造SI小 = r.PA60本厂铸造SI小;
                            x.P61送炼钢厂SI大 = r.PA61送炼钢厂SI大;
                            x.P62送炼钢厂SI中 = r.PA62送炼钢厂SI中;
                            x.P63送炼钢厂SI小 = r.PA63送炼钢厂SI小;
                            n.Add(x);
                        }
                        return n;
                    }
                    else
                    {
                        return null;
                    }
                }
            }


        }

        partial class 技术日报Row
        {
            
            public double? PA03合格铁
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P03合格铁Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P03合格铁”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P03合格铁Column] = value;
                }
            }

            
            public double? PA04炼钢铁
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P04炼钢铁Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P04炼钢铁”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P04炼钢铁Column] = value;
                }
            }

            
            public double? PA05铸造铁
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P05铸造铁Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P05铸造铁”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P05铸造铁Column] = value;
                }
            }

            
            public double? PA06号外铁
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P06号外铁Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P06号外铁”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P06号外铁Column] = value;
                }
            }

            
            public double? PA07合格率
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P07合格率Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P07合格率”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P07合格率Column] = value;
                }
            }

            
            public double? PA08高炉利用系数
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P08高炉利用系数Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P08高炉利用系数”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P08高炉利用系数Column] = value;
                }
            }

            
            public double? PA09高炉实物系数
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P09高炉实物系数Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P09高炉实物系数”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P09高炉实物系数Column] = value;
                }
            }

            
            public double? PA10原料矿合计总耗
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P10原料矿合计总耗Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P10原料矿合计总耗”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P10原料矿合计总耗Column] = value;
                }
            }

            
            public double? PA11原料矿合计单耗
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P11原料矿合计单耗Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P11原料矿合计单耗”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P11原料矿合计单耗Column] = value;
                }
            }

            
            public double? PA12原料矿机烧
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P12原料矿机烧Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P12原料矿机烧”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P12原料矿机烧Column] = value;
                }
            }

            
            public double? PA13原料矿竖炉球
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P13原料矿竖炉球Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P13原料矿竖炉球”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P13原料矿竖炉球Column] = value;
                }
            }

            
            public double? PA14原料矿CT
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P14原料矿CTColumn]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P14原料矿CT”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P14原料矿CTColumn] = value;
                }
            }

            
            public double? PA15原料矿其它熟料
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P15原料矿其它熟料Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P15原料矿其它熟料”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P15原料矿其它熟料Column] = value;
                }
            }

            
            public double? PA16原料矿本溪矿
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P16原料矿本溪矿Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P16原料矿本溪矿”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P16原料矿本溪矿Column] = value;
                }
            }

            
            public double? PA17原料矿其它生料
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P17原料矿其它生料Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P17原料矿其它生料”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P17原料矿其它生料Column] = value;
                }
            }

            
            public double? PA18废铁总耗
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P18废铁总耗Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P18废铁总耗”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P18废铁总耗Column] = value;
                }
            }

            
            public double? PA19废铁单耗
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P19废铁单耗Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P19废铁单耗”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P19废铁单耗Column] = value;
                }
            }

            
            public double? PA20回收率
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P20回收率Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P20回收率”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P20回收率Column] = value;
                }
            }

            
            public double? PA21熟料比
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P21熟料比Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P21熟料比”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P21熟料比Column] = value;
                }
            }

            
            public double? PA22平均风温
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P22平均风温Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P22平均风温”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P22平均风温Column] = value;
                }
            }

            
            public double? PA23炉顶温度
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P23炉顶温度Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P23炉顶温度”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P23炉顶温度Column] = value;
                }
            }

            
            public double? PA24热风压力
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P24热风压力Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P24热风压力”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P24热风压力Column] = value;
                }
            }

            
            public double? PA25炉顶压力
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P25炉顶压力Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P25炉顶压力”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P25炉顶压力Column] = value;
                }
            }

            
            public double? PA26富氧率
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P26富氧率Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P26富氧率”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P26富氧率Column] = value;
                }
            }

            
            public double? PA27入炉焦炭总耗
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P27入炉焦炭总耗Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P27入炉焦炭总耗”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P27入炉焦炭总耗Column] = value;
                }
            }

            
            public double? PA28入炉焦炭单耗
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P28入炉焦炭单耗Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P28入炉焦炭单耗”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P28入炉焦炭单耗Column] = value;
                }
            }

            
            public double? PA29煤粉总耗
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P29煤粉总耗Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P29煤粉总耗”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P29煤粉总耗Column] = value;
                }
            }

            
            public double? PA30煤粉单耗
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P30煤粉单耗Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P30煤粉单耗”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P30煤粉单耗Column] = value;
                }
            }

            
            public double? PA31焦丁总耗
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P31焦丁总耗Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P31焦丁总耗”的值为 DBNull。", e);
                    }
                }

            }

            
            public double? PA32焦丁单耗
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P32焦丁单耗Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P32焦丁单耗”的值为 DBNull。", e);
                    }
                }

            }

            
            public double? PA33综合焦炭总耗
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P33综合焦炭总耗Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P33综合焦炭总耗”的值为 DBNull。", e);
                    }
                }
            }

            
            public double? PA34综合焦炭单耗
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P34综合焦炭单耗Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P34综合焦炭单耗”的值为 DBNull。", e);
                    }
                }
            }

            
            public double? PA35综合折算焦比
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P35综合折算焦比Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P35综合折算焦比”的值为 DBNull。", e);
                    }
                }
            }

            
            public double? PA36冶炼强度
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P36冶炼强度Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P36冶炼强度”的值为 DBNull。", e);
                    }
                }
            }

            
            public double? PA37焦炭负荷
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P37焦炭负荷Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P37焦炭负荷”的值为 DBNull。", e);
                    }
                }
            }

            
            public double? PA38干毛焦
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P38干毛焦Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P38干毛焦”的值为 DBNull。", e);
                    }
                }

            }

            
            public double? PA39炼钢铁SI
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P39炼钢铁SIColumn]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P39炼钢铁SI”的值为 DBNull。", e);
                    }
                }

            }

            
            public double? PA40炼钢铁MN
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P40炼钢铁MNColumn]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P40炼钢铁MN”的值为 DBNull。", e);
                    }
                }

            }

            
            public double? PA41炼钢铁P
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P41炼钢铁PColumn]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P41炼钢铁P”的值为 DBNull。", e);
                    }
                }

            }

            
            public double? PA42炼钢铁S
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P42炼钢铁SColumn]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P42炼钢铁S”的值为 DBNull。", e);
                    }
                }

            }

            
            public double? PA43铸造铁SI
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P43铸造铁SIColumn]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P43铸造铁SI”的值为 DBNull。", e);
                    }
                }

            }

            
            public double? PA44铸造铁MN
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P44铸造铁MNColumn]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P44铸造铁MN”的值为 DBNull。", e);
                    }
                }

            }

            
            public double? PA45铸造铁P
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P45铸造铁PColumn]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P45铸造铁P”的值为 DBNull。", e);
                    }
                }

            }

            
            public double? PA46铸造铁S
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P46铸造铁SColumn]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P46铸造铁S”的值为 DBNull。", e);
                    }
                }

            }

            
            public double? PA47炉渣碱度
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P47炉渣碱度Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P47炉渣碱度”的值为 DBNull。", e);
                    }
                }

            }

            
            public double? PA48休风情况
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P48休风情况Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P48休风情况”的值为 DBNull。", e);
                    }
                }

            }

            
            public double? PA49慢风
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P49慢风Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P49慢风”的值为 DBNull。", e);
                    }
                }

            }

            
            public double? PA50坐料次数
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P50坐料次数Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P50坐料次数”的值为 DBNull。", e);
                    }
                }

            }

            
            public double? PA51悬料次数
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P51悬料次数Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P51悬料次数”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P51悬料次数Column] = value;
                }
            }

            
            public double? PA52崩料次数
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P52崩料次数Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P52崩料次数”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P52崩料次数Column] = value;
                }
            }

            
            public double? PA53风口损坏数大
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P53风口损坏数大Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P53风口损坏数大”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P53风口损坏数大Column] = value;
                }
            }

            
            public double? PA54风口损坏数中
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P54风口损坏数中Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P54风口损坏数中”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P54风口损坏数中Column] = value;
                }
            }

            
            public double? PA55风口损坏数小
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P55风口损坏数小Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P55风口损坏数小”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P55风口损坏数小Column] = value;
                }
            }

            
            public double? PA56渣口损坏数大
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P56渣口损坏数大Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P56渣口损坏数大”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P56渣口损坏数大Column] = value;
                }
            }

            
            public double? PA57渣口损坏数中
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P57渣口损坏数中Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P57渣口损坏数中”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P57渣口损坏数中Column] = value;
                }
            }

            
            public double? PA58渣口损坏数小
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P58渣口损坏数小Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P58渣口损坏数小”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P58渣口损坏数小Column] = value;
                }
            }

            
            public double? PA59本厂铸造SI大
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P59本厂铸造SI大Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P59本厂铸造SI大”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P59本厂铸造SI大Column] = value;
                }
            }

            
            public double? PA60本厂铸造SI小
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P60本厂铸造SI小Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P60本厂铸造SI小”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P60本厂铸造SI小Column] = value;
                }
            }

            
            public double? PA61送炼钢厂SI大
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P61送炼钢厂SI大Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P61送炼钢厂SI大”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P61送炼钢厂SI大Column] = value;
                }
            }

            
            public double? PA62送炼钢厂SI中
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P62送炼钢厂SI中Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P62送炼钢厂SI中”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P62送炼钢厂SI中Column] = value;
                }
            }

            
            public double? PA63送炼钢厂SI小
            {
                get
                {
                    try
                    {
                        return ((double)(this[this.table技术日报.P63送炼钢厂SI小Column]));
                    }
                    catch 
                    {
                        return null; //throw new System.Data.StrongTypingException("表“技术日报”中列“P63送炼钢厂SI小”的值为 DBNull。", e);
                    }
                }
                set
                {
                    this[this.table技术日报.P63送炼钢厂SI小Column] = value;
                }
            }
        }
    }
}
