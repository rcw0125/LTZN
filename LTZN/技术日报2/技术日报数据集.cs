namespace LTZN.技术日报 {


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
                       
                            x.P03合格铁 = r.IsP03合格铁Null() ? 0 : (double?)r.P03合格铁;
                            x.P04炼钢铁 = r.IsP04炼钢铁Null() ? 0 : (double?)r.P04炼钢铁;
                            x.P05铸造铁 = r.IsP05铸造铁Null() ? 0 : (double?)r.P05铸造铁;
                            x.P06号外铁 = r.IsP06号外铁Null() ? 0 : (double?)r.P06号外铁;
                            x.P07合格率 = r.IsP07合格率Null() ? 0 : (double?)r.P07合格率;
                            x.P08高炉利用系数 = r.IsP08高炉利用系数Null() ? 0 : (double?)r.P08高炉利用系数;
                            x.P09高炉实物系数 = r.IsP09高炉实物系数Null() ? 0 : (double?)r.P09高炉实物系数;
                            x.P10原料矿合计总耗 = r.IsP10原料矿合计总耗Null() ? 0 : (double?)r.P10原料矿合计总耗;
                            x.P11原料矿合计单耗 = r.IsP11原料矿合计单耗Null() ? 0 : (double?)r.P11原料矿合计单耗;
                            x.P12原料矿机烧 = r.IsP12原料矿机烧Null() ? 0 : (double?)r.P12原料矿机烧;
                            x.P13原料矿竖炉球 = r.IsP13原料矿竖炉球Null() ? 0 : (double?)r.P13原料矿竖炉球;
                            x.P14原料矿CT = r.IsP14原料矿CTNull() ? 0 : (double?)r.P14原料矿CT;
                            x.P15原料矿其它熟料 = r.IsP15原料矿其它熟料Null() ? 0 : (double?)r.P15原料矿其它熟料;
                            x.P16原料矿本溪矿 = r.IsP16原料矿本溪矿Null() ? 0 : (double?)r.P16原料矿本溪矿;
                            x.P17原料矿其它生料 = r.IsP17原料矿其它生料Null() ? 0 : (double?)r.P17原料矿其它生料;
                            x.P18废铁总耗 = r.IsP18废铁总耗Null() ? 0 : (double?)r.P18废铁总耗;
                            x.P19废铁单耗 = r.IsP19废铁单耗Null() ? 0 : (double?)r.P19废铁单耗;
                            x.P20回收率 = r.IsP20回收率Null() ? 0 : (double?)r.P20回收率;
                            x.P21熟料比 = r.IsP21熟料比Null() ? 0 : (double?)r.P21熟料比;
                            x.P22平均风温 = r.IsP22平均风温Null() ? 0 : (double?)r.P22平均风温;
                            x.P23炉顶温度 = r.IsP23炉顶温度Null() ? 0 : (double?)r.P23炉顶温度;
                            x.P24热风压力 = r.IsP24热风压力Null() ? 0 : (double?)r.P24热风压力;
                            x.P25炉顶压力 = r.IsP25炉顶压力Null() ? 0 : (double?)r.P25炉顶压力;
                            x.P26富氧率 = r.IsP26富氧率Null() ? 0 : (double?)r.P26富氧率;
                            x.P27入炉焦炭总耗 = r.IsP27入炉焦炭总耗Null() ? 0 : (double?)r.P27入炉焦炭总耗;
                            x.P28入炉焦炭单耗 = r.IsP28入炉焦炭单耗Null() ? 0 : (double?)r.P28入炉焦炭单耗;
                            x.P29煤粉总耗 = r.IsP29煤粉总耗Null() ? 0 : (double?)r.P29煤粉总耗;
                            x.P30煤粉单耗 = r.IsP30煤粉单耗Null() ? 0 : (double?)r.P30煤粉单耗;
                            x.P31焦丁总耗 = r.IsP31焦丁总耗Null() ? 0 : (double?)r.P31焦丁总耗;
                            x.P32焦丁单耗 = r.IsP32焦丁单耗Null() ? 0 : (double?)r.P32焦丁单耗;
                            x.P33综合焦炭总耗 = r.IsP33综合焦炭总耗Null() ? 0 : (double?)r.P33综合焦炭总耗;
                            x.P34综合焦炭单耗 = r.IsP34综合焦炭单耗Null() ? 0 : (double?)r.P34综合焦炭单耗;
                            x.P35综合折算焦比 = r.IsP35综合折算焦比Null() ? 0 : (double?)r.P35综合折算焦比;
                            x.P36冶炼强度 = r.IsP36冶炼强度Null() ? 0 : (double?)r.P36冶炼强度;
                            x.P37焦炭负荷 = r.IsP37焦炭负荷Null() ? 0 : (double?)r.P37焦炭负荷;
                            x.P38干毛焦 = r.IsP38干毛焦Null() ? 0 : (double?)r.P38干毛焦;
                            x.P39炼钢铁SI = r.IsP39炼钢铁SINull() ? 0 : (double?)r.P39炼钢铁SI;
                            x.P40炼钢铁MN = r.IsP40炼钢铁MNNull() ? 0 : (double?)r.P40炼钢铁MN;
                            x.P41炼钢铁P = r.IsP41炼钢铁PNull() ? 0 : (double?)r.P41炼钢铁P;
                            x.P42炼钢铁S = r.IsP42炼钢铁SNull() ? 0 : (double?)r.P42炼钢铁S;
                            x.P43铸造铁SI = r.IsP43铸造铁SINull() ? 0 : (double?)r.P43铸造铁SI;
                            x.P44铸造铁MN = r.IsP44铸造铁MNNull() ? 0 : (double?)r.P44铸造铁MN;
                            x.P45铸造铁P = r.IsP45铸造铁PNull() ? 0 : (double?)r.P45铸造铁P;
                            x.P46铸造铁S = r.IsP46铸造铁SNull() ? 0 : (double?)r.P46铸造铁S;
                            x.P47炉渣碱度 = r.IsP47炉渣碱度Null() ? 0 : (double?)r.P47炉渣碱度;
                            x.P48休风情况 = r.IsP48休风情况Null() ? 0 : (double?)r.P48休风情况;
                            x.P49慢风 = r.IsP49慢风Null() ? 0 : (double?)r.P49慢风;
                            x.P50坐料次数 = r.IsP50坐料次数Null() ? 0 : (double?)r.P50坐料次数;
                            x.P51悬料次数 = r.IsP51悬料次数Null() ? 0 : (double?)r.P51悬料次数;
                            x.P52崩料次数 = r.IsP52崩料次数Null() ? 0 : (double?)r.P52崩料次数;
                            x.P53风口损坏数大 = r.IsP53风口损坏数大Null() ? 0 : (double?)r.P53风口损坏数大;
                            x.P54风口损坏数中 = r.IsP54风口损坏数中Null() ? 0 : (double?)r.P54风口损坏数中;
                            x.P55风口损坏数小 = r.IsP55风口损坏数小Null() ? 0 : (double?)r.P55风口损坏数小;
                            x.P56渣口损坏数大 = r.IsP56渣口损坏数大Null() ? 0 : (double?)r.P56渣口损坏数大;
                            x.P57渣口损坏数中 = r.IsP57渣口损坏数中Null() ? 0 : (double?)r.P57渣口损坏数中;
                            x.P58渣口损坏数小 = r.IsP58渣口损坏数小Null() ? 0 : (double?)r.P58渣口损坏数小;
                            x.P59本厂铸造SI大 = r.IsP59本厂铸造SI大Null() ? 0 : (double?)r.P59本厂铸造SI大;
                            x.P60本厂铸造SI小 = r.IsP60本厂铸造SI小Null() ? 0 : (double?)r.P60本厂铸造SI小;
                            x.P61送炼钢厂SI大 = r.IsP61送炼钢厂SI大Null() ? 0 : (double?)r.P61送炼钢厂SI大;
                            x.P62送炼钢厂SI中 = r.IsP62送炼钢厂SI中Null() ? 0 : (double?)r.P62送炼钢厂SI中;
                            x.P63送炼钢厂SI小 = r.IsP63送炼钢厂SI小Null() ? 0 : (double?)r.P63送炼钢厂SI小;
                            x.P64折算产量 = r.IsP64折算产量Null() ? 0 : (double?)r.P64折算产量;
                            x.P生成标志 = r.IsP生成标志Null() ? 0 : (double?)r.P生成标志;
                            n.Add(x);
                        }
                        return n;
                    }
                    else
                    {
                        return null;
                    }
                }
                set
                {
                    foreach (技术日报内容项 x in value)
                    {
                        技术日报Row r = this.FindByP日期P01单位P02项目(value.riqi,x.P01单位,x.P02项目);
                        if (x.p[0])
                        r["P03合格铁"] =ConvertToDBObject(x.P03合格铁);
                        if (x.p[1])
                        r["P04炼钢铁"] = ConvertToDBObject(x.P04炼钢铁);
                        if (x.p[2])
                        r["P05铸造铁"] = ConvertToDBObject(x.P05铸造铁);
                        if (x.p[3])
                        r["P06号外铁"] = ConvertToDBObject(x.P06号外铁);
                        if (x.p[4])
                        r["P07合格率"] = ConvertToDBObject(x.P07合格率);
                        if (x.p[5])
                        r["P08高炉利用系数"] = ConvertToDBObject(x.P08高炉利用系数);
                        if (x.p[6])
                        r["P09高炉实物系数"] = ConvertToDBObject(x.P09高炉实物系数);
                        if (x.p[7])
                        r["P10原料矿合计总耗"] = ConvertToDBObject(x.P10原料矿合计总耗);
                        if (x.p[8])
                        r["P11原料矿合计单耗"] = ConvertToDBObject(x.P11原料矿合计单耗);
                        if (x.p[9])
                        r["P12原料矿机烧"] = ConvertToDBObject(x.P12原料矿机烧);
                        if (x.p[10])
                        r["P13原料矿竖炉球"] = ConvertToDBObject(x.P13原料矿竖炉球);
                        if (x.p[11])
                        r["P14原料矿CT"] = ConvertToDBObject(x.P14原料矿CT);
                        if (x.p[12])
                        r["P15原料矿其它熟料"] = ConvertToDBObject(x.P15原料矿其它熟料);
                        if (x.p[13])
                        r["P16原料矿本溪矿"] = ConvertToDBObject(x.P16原料矿本溪矿);
                        if (x.p[14])
                        r["P17原料矿其它生料"] = ConvertToDBObject(x.P17原料矿其它生料);
                        if (x.p[15])
                        r["P18废铁总耗"] = ConvertToDBObject(x.P18废铁总耗);
                        if (x.p[16])
                        r["P19废铁单耗"] = ConvertToDBObject(x.P19废铁单耗);
                        if (x.p[17])
                        r["P20回收率"] = ConvertToDBObject(x.P20回收率);
                        if (x.p[18])
                        r["P21熟料比"] = ConvertToDBObject(x.P21熟料比);
                        if (x.p[19])
                        r["P22平均风温"] = ConvertToDBObject(x.P22平均风温);
                        if (x.p[20])
                        r["P23炉顶温度"] = ConvertToDBObject(x.P23炉顶温度);
                        if (x.p[21])
                        r["P24热风压力"] = ConvertToDBObject(x.P24热风压力);
                        if (x.p[22])
                        r["P25炉顶压力"] = ConvertToDBObject(x.P25炉顶压力);
                        if (x.p[23])
                        r["P26富氧率"] = ConvertToDBObject(x.P26富氧率);
                        if (x.p[24])
                        r["P27入炉焦炭总耗"] = ConvertToDBObject(x.P27入炉焦炭总耗);
                        if (x.p[25])
                        r["P28入炉焦炭单耗"] = ConvertToDBObject(x.P28入炉焦炭单耗);
                        if (x.p[26])
                        r["P29煤粉总耗"] = ConvertToDBObject(x.P29煤粉总耗);
                        if (x.p[27])
                        r["P30煤粉单耗"] = ConvertToDBObject(x.P30煤粉单耗);
                        if (x.p[28])
                        r["P31焦丁总耗"] = ConvertToDBObject(x.P31焦丁总耗);
                        if (x.p[29])
                        r["P32焦丁单耗"] = ConvertToDBObject(x.P32焦丁单耗);
                        if (x.p[30])
                        r["P33综合焦炭总耗"] = ConvertToDBObject(x.P33综合焦炭总耗);
                        if (x.p[31])
                        r["P34综合焦炭单耗"] = ConvertToDBObject(x.P34综合焦炭单耗);
                        if (x.p[32])
                        r["P35综合折算焦比"] = ConvertToDBObject(x.P35综合折算焦比);
                        if (x.p[33])
                        r["P36冶炼强度"] = ConvertToDBObject(x.P36冶炼强度);
                        if (x.p[34])
                        r["P37焦炭负荷"] = ConvertToDBObject(x.P37焦炭负荷);
                        if (x.p[35])
                        r["P38干毛焦"] = ConvertToDBObject(x.P38干毛焦);
                        if (x.p[36])
                        r["P39炼钢铁SI"] = ConvertToDBObject(x.P39炼钢铁SI);
                        if (x.p[37])
                        r["P40炼钢铁MN"] = ConvertToDBObject(x.P40炼钢铁MN);
                        if (x.p[38])
                        r["P41炼钢铁P"] = ConvertToDBObject(x.P41炼钢铁P);
                        if (x.p[39])
                        r["P42炼钢铁S"] = ConvertToDBObject(x.P42炼钢铁S);
                        if (x.p[40])
                        r["P43铸造铁SI"] = ConvertToDBObject(x.P43铸造铁SI);
                        if (x.p[41])
                        r["P44铸造铁MN"] = ConvertToDBObject(x.P44铸造铁MN);
                        if (x.p[42])
                        r["P45铸造铁P"] = ConvertToDBObject(x.P45铸造铁P);
                        if (x.p[43])
                        r["P46铸造铁S"] = ConvertToDBObject(x.P46铸造铁S);
                        if (x.p[44])
                        r["P47炉渣碱度"] = ConvertToDBObject(x.P47炉渣碱度);
                        if (x.p[45])
                        r["P48休风情况"] = ConvertToDBObject(x.P48休风情况);
                        if (x.p[46])
                        r["P49慢风"] = ConvertToDBObject(x.P49慢风);
                        if (x.p[47])
                        r["P50坐料次数"] = ConvertToDBObject(x.P50坐料次数);
                        if (x.p[48])
                        r["P51悬料次数"] = ConvertToDBObject(x.P51悬料次数);
                        if (x.p[49])
                        r["P52崩料次数"] = ConvertToDBObject(x.P52崩料次数);
                        if (x.p[50])
                        r["P53风口损坏数大"] = ConvertToDBObject(x.P53风口损坏数大);
                        if (x.p[51])
                        r["P54风口损坏数中"] = ConvertToDBObject(x.P54风口损坏数中);
                        if (x.p[52])
                        r["P55风口损坏数小"] = ConvertToDBObject(x.P55风口损坏数小);
                        if (x.p[53])
                        r["P56渣口损坏数大"] = ConvertToDBObject(x.P56渣口损坏数大);
                        if (x.p[54])
                        r["P57渣口损坏数中"] = ConvertToDBObject(x.P57渣口损坏数中);
                        if (x.p[55])
                        r["P58渣口损坏数小"] = ConvertToDBObject(x.P58渣口损坏数小);
                        if (x.p[56])
                        r["P59本厂铸造SI大"] = ConvertToDBObject(x.P59本厂铸造SI大);
                        if (x.p[57])
                        r["P60本厂铸造SI小"] = ConvertToDBObject(x.P60本厂铸造SI小);
                        if (x.p[58])
                        r["P61送炼钢厂SI大"] = ConvertToDBObject(x.P61送炼钢厂SI大);
                        if (x.p[59])
                        r["P62送炼钢厂SI中"] = ConvertToDBObject(x.P62送炼钢厂SI中);
                        if (x.p[60])
                        r["P63送炼钢厂SI小"] = ConvertToDBObject(x.P63送炼钢厂SI小);
                       if (x.p[61])
                        r["P64折算产量"] = ConvertToDBObject(x.P64折算产量);
                       if (x.p[62])
                        r["P生成标志"] = ConvertToDBObject(x.P生成标志);
                    }
                }
            }

            private object ConvertToDBObject(double? d)
            {
                try
                {
                    if (d.HasValue)
                        if (double.IsNaN(d.Value))
                            return 0;
                        else
                            return System.Convert.ToDecimal(d.Value);
                    else
                        return System.DBNull.Value;
                }
                catch { }
                return System.DBNull.Value;
            }

        }
    }
}
