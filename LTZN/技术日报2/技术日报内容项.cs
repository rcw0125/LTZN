using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;

namespace LTZN.技术日报
{
    [DataObjectAttribute()]
    public class 技术日报内容项
    {
        #region 63个属性

        public bool[] p = new bool[63];

        private 技术日报内容项数据 currentData;

        [Bindable(true)]
        public string P01单位
        {
            get
            {
                return this.currentData.p01单位;
            }
            set
            {
                this.currentData.p01单位 = value;
            }
        }

        [Bindable(true)]
        public string P02项目
        {
            get
            {
                return this.currentData.p02项目;
            }
            set
            {
                this.currentData.p02项目 = value;
            }
        }

        [Bindable(true)]
        public double? P03合格铁
        {
            get
            {
                return this.currentData.p03合格铁;
            }
            set
            {
                this.currentData.p03合格铁 = value;
            }
        }

        [Bindable(true)]
        public double? P04炼钢铁
        {
            get
            {
                return this.currentData.p04炼钢铁;
            }
            set
            {
                this.currentData.p04炼钢铁 = value;
            }
        }

        [Bindable(true)]
        public double? P05铸造铁
        {
            get
            {
                return this.currentData.p05铸造铁;
            }
            set
            {
                this.currentData.p05铸造铁 = value;
            }
        }

        [Bindable(true)]
        public double? P06号外铁
        {
            get
            {
                return this.currentData.p06号外铁;
            }
            set
            {
                this.currentData.p06号外铁 = value;
            }
        }

        [Bindable(true)]
        public double? P07合格率
        {
            get
            {
                return this.currentData.p07合格率;
            }
            set
            {
                this.currentData.p07合格率 = value;
            }
        }

        [Bindable(true)]
        public double? P08高炉利用系数
        {
            get
            {
                return this.currentData.p08高炉利用系数;
            }
            set
            {
                this.currentData.p08高炉利用系数 = value;
            }
        }

        [Bindable(true)]
        public double? P09高炉实物系数
        {
            get
            {
                return this.currentData.p09高炉实物系数;
            }
            set
            {
                this.currentData.p09高炉实物系数 = value;
            }
        }

        [Bindable(true)]
        public double? P10原料矿合计总耗
        {
            get
            {
                return this.currentData.p10原料矿合计总耗;
            }
            set
            {
                this.currentData.p10原料矿合计总耗 = value;
            }
        }

        [Bindable(true)]
        public double? P11原料矿合计单耗
        {
            get
            {
                return this.currentData.p11原料矿合计单耗;
            }
            set
            {
                this.currentData.p11原料矿合计单耗 = value;
            }
        }

        [Bindable(true)]
        public double? P12原料矿机烧
        {
            get
            {
                return this.currentData.p12原料矿机烧;
            }
            set
            {
                this.currentData.p12原料矿机烧 = value;
            }
        }

        [Bindable(true)]
        public double? P13原料矿竖炉球
        {
            get
            {
                return this.currentData.p13原料矿竖炉球;
            }
            set
            {
                this.currentData.p13原料矿竖炉球 = value;
            }
        }

        [Bindable(true)]
        public double? P14原料矿CT
        {
            get
            {
                return this.currentData.p14原料矿CT;
            }
            set
            {
                this.currentData.p14原料矿CT = value;
            }
        }

        [Bindable(true)]
        public double? P15原料矿其它熟料
        {
            get
            {
                return this.currentData.p15原料矿其它熟料;
            }
            set
            {
                this.currentData.p15原料矿其它熟料 = value;
            }
        }

        [Bindable(true)]
        public double? P16原料矿本溪矿
        {
            get
            {
                return this.currentData.p16原料矿本溪矿;
            }
            set
            {
                this.currentData.p16原料矿本溪矿 = value;
            }
        }

        [Bindable(true)]
        public double? P17原料矿其它生料
        {
            get
            {
                return this.currentData.p17原料矿其它生料;
            }
            set
            {
                this.currentData.p17原料矿其它生料 = value;
            }
        }

        [Bindable(true)]
        public double? P18废铁总耗
        {
            get
            {
                return this.currentData.p18废铁总耗;
            }
            set
            {
                this.currentData.p18废铁总耗 = value;
            }
        }

        [Bindable(true)]
        public double? P19废铁单耗
        {
            get
            {
                return this.currentData.p19废铁单耗;
            }
            set
            {
                this.currentData.p19废铁单耗 = value;
            }
        }

        [Bindable(true)]
        public double? P20回收率
        {
            get
            {
                return this.currentData.p20回收率;
            }
            set
            {
                this.currentData.p20回收率 = value;
            }
        }

        [Bindable(true)]
        public double? P21熟料比
        {
            get
            {
                return this.currentData.p21熟料比;
            }
            set
            {
                this.currentData.p21熟料比 = value;
            }
        }

        [Bindable(true)]
        public double? P22平均风温
        {
            get
            {
                return this.currentData.p22平均风温;
            }
            set
            {
                this.currentData.p22平均风温 = value;
            }
        }

        [Bindable(true)]
        public double? P23炉顶温度
        {
            get
            {
                return this.currentData.p23炉顶温度;
            }
            set
            {
                this.currentData.p23炉顶温度 = value;
            }
        }

        [Bindable(true)]
        public double? P24热风压力
        {
            get
            {
                return this.currentData.p24热风压力;
            }
            set
            {
                this.currentData.p24热风压力 = value;
            }
        }

        [Bindable(true)]
        public double? P25炉顶压力
        {
            get
            {
                return this.currentData.p25炉顶压力;
            }
            set
            {
                this.currentData.p25炉顶压力 = value;
            }
        }

        [Bindable(true)]
        public double? P26富氧率
        {
            get
            {
                return this.currentData.p26富氧率;
            }
            set
            {
                this.currentData.p26富氧率 = value;
            }
        }

        [Bindable(true)]
        public double? P27入炉焦炭总耗
        {
            get
            {
                return this.currentData.p27入炉焦炭总耗;
            }
            set
            {
                this.currentData.p27入炉焦炭总耗 = value;
            }
        }

        [Bindable(true)]
        public double? P28入炉焦炭单耗
        {
            get
            {
                return this.currentData.p28入炉焦炭单耗;
            }
            set
            {
                this.currentData.p28入炉焦炭单耗 = value;
            }
        }

        [Bindable(true)]
        public double? P29煤粉总耗
        {
            get
            {
                return this.currentData.p29煤粉总耗;
            }
            set
            {
                this.currentData.p29煤粉总耗 = value;
            }
        }

        [Bindable(true)]
        public double? P30煤粉单耗
        {
            get
            {
                return this.currentData.p30煤粉单耗;
            }
            set
            {
                this.currentData.p30煤粉单耗 = value;
            }
        }

        [Bindable(true)]
        public double? P31焦丁总耗
        {
            get
            {
                return this.currentData.p31焦丁总耗;
            }
            set
            {
                this.currentData.p31焦丁总耗 = value;
            }
        }

        [Bindable(true)]
        public double? P32焦丁单耗
        {
            get
            {
                return this.currentData.p32焦丁单耗;
            }
            set
            {
                this.currentData.p32焦丁单耗 = value;
            }
        }

        [Bindable(true)]
        public double? P33综合焦炭总耗
        {
            get
            {
                return this.currentData.p33综合焦炭总耗;
            }
            set
            {
                this.currentData.p33综合焦炭总耗 = value;
            }
        }

        [Bindable(true)]
        public double? P34综合焦炭单耗
        {
            get
            {
                return this.currentData.p34综合焦炭单耗;
            }
            set
            {
                this.currentData.p34综合焦炭单耗 = value;
            }
        }

        [Bindable(true)]
        public double? P35综合折算焦比
        {
            get
            {
                return this.currentData.p35综合折算焦比;
            }
            set
            {
                this.currentData.p35综合折算焦比 = value;
            }
        }

        [Bindable(true)]
        public double? P36冶炼强度
        {
            get
            {
                return this.currentData.p36冶炼强度;
            }
            set
            {
                this.currentData.p36冶炼强度 = value;
            }
        }

        [Bindable(true)]
        public double? P37焦炭负荷
        {
            get
            {
                return this.currentData.p37焦炭负荷;
            }
            set
            {
                this.currentData.p37焦炭负荷 = value;
            }
        }

        [Bindable(true)]
        public double? P38干毛焦
        {
            get
            {
                return this.currentData.p38干毛焦;
            }
            set
            {
                this.currentData.p38干毛焦 = value;
            }
        }

        [Bindable(true)]
        public double? P39炼钢铁SI
        {
            get
            {
                return this.currentData.p39炼钢铁SI;
            }
            set
            {
                this.currentData.p39炼钢铁SI = value;
            }
        }

        [Bindable(true)]
        public double? P40炼钢铁MN
        {
            get
            {
                return this.currentData.p40炼钢铁MN;
            }
            set
            {
                this.currentData.p40炼钢铁MN = value;
            }
        }
        [Bindable(true)]

        public double? P41炼钢铁P
        {
            get
            {
                return this.currentData.p41炼钢铁P;
            }
            set
            {
                this.currentData.p41炼钢铁P = value;
            }
        }

        [Bindable(true)]
        public double? P42炼钢铁S
        {
            get
            {
                return this.currentData.p42炼钢铁S;
            }
            set
            {
                this.currentData.p42炼钢铁S = value;
            }
        }

        [Bindable(true)]
        public double? P43铸造铁SI
        {
            get
            {
                return this.currentData.p43铸造铁SI;
            }
            set
            {
                this.currentData.p43铸造铁SI = value;
            }
        }

        [Bindable(true)]
        public double? P44铸造铁MN
        {
            get
            {
                return this.currentData.p44铸造铁MN;
            }
            set
            {
                this.currentData.p44铸造铁MN = value;
            }
        }

        [Bindable(true)]
        public double? P45铸造铁P
        {
            get
            {
                return this.currentData.p45铸造铁P;
            }
            set
            {
                this.currentData.p45铸造铁P = value;
            }
        }

        [Bindable(true)]
        public double? P46铸造铁S
        {
            get
            {
                return this.currentData.p46铸造铁S;
            }
            set
            {
                this.currentData.p46铸造铁S = value;
            }
        }

        [Bindable(true)]
        public double? P47炉渣碱度
        {
            get
            {
                return this.currentData.p47炉渣碱度;
            }
            set
            {
                this.currentData.p47炉渣碱度 = value;
            }
        }

        [Bindable(true)]
        public double? P48休风情况
        {
            get
            {
                return this.currentData.p48休风情况;
            }
            set
            {
                this.currentData.p48休风情况 = value;
            }
        }

        [Bindable(true)]
        public double? P49慢风
        {
            get
            {
                return this.currentData.p49慢风;
            }
            set
            {
                this.currentData.p49慢风 = value;
            }
        }

        [Bindable(true)]
        public double? P50坐料次数
        {
            get
            {
                return this.currentData.p50坐料次数;
            }
            set
            {
                this.currentData.p50坐料次数 = value;
            }
        }

        [Bindable(true)]
        public double? P51悬料次数
        {
            get
            {
                return this.currentData.p51悬料次数;
            }
            set
            {
                this.currentData.p51悬料次数 = value;
            }
        }

        [Bindable(true)]
        public double? P52崩料次数
        {
            get
            {
                return this.currentData.p52崩料次数;
            }
            set
            {
                this.currentData.p52崩料次数 = value;
            }
        }

        [Bindable(true)]
        public double? P53风口损坏数大
        {
            get
            {
                return this.currentData.p53风口损坏数大;
            }
            set
            {
                this.currentData.p53风口损坏数大 = value;
            }
        }

        [Bindable(true)]
        public double? P54风口损坏数中
        {
            get
            {
                return this.currentData.p54风口损坏数中;
            }
            set
            {
                this.currentData.p54风口损坏数中 = value;
            }
        }

        [Bindable(true)]
        public double? P55风口损坏数小
        {
            get
            {
                return this.currentData.p55风口损坏数小;
            }
            set
            {
                this.currentData.p55风口损坏数小 = value;
            }
        }

        [Bindable(true)]
        public double? P56渣口损坏数大
        {
            get
            {
                return this.currentData.p56渣口损坏数大;
            }
            set
            {
                this.currentData.p56渣口损坏数大 = value;
            }
        }

        [Bindable(true)]
        public double? P57渣口损坏数中
        {
            get
            {
                return this.currentData.p57渣口损坏数中;
            }
            set
            {
                this.currentData.p57渣口损坏数中 = value;
            }
        }

        [Bindable(true)]
        public double? P58渣口损坏数小
        {
            get
            {
                return this.currentData.p58渣口损坏数小;
            }
            set
            {
                this.currentData.p58渣口损坏数小 = value;
            }
        }

        [Bindable(true)]
        public double? P59本厂铸造SI大
        {
            get
            {
                return this.currentData.p59本厂铸造SI大;
            }
            set
            {
                this.currentData.p59本厂铸造SI大 = value;
            }
        }

        [Bindable(true)]
        public double? P60本厂铸造SI小
        {
            get
            {
                return this.currentData.p60本厂铸造SI小;
            }
            set
            {
                this.currentData.p60本厂铸造SI小 = value;
            }
        }

        [Bindable(true)]
        public double? P61送炼钢厂SI大
        {
            get
            {
                return this.currentData.p61送炼钢厂SI大;
            }
            set
            {
                this.currentData.p61送炼钢厂SI大 = value;
            }
        }

        [Bindable(true)]
        public double? P62送炼钢厂SI中
        {
            get
            {
                return this.currentData.p62送炼钢厂SI中;
            }
            set
            {
                this.currentData.p62送炼钢厂SI中 = value;
            }
        }

        [Bindable(true)]
        public double? P63送炼钢厂SI小
        {
            get
            {
                return this.currentData.p63送炼钢厂SI小;
            }
            set
            {
                this.currentData.p63送炼钢厂SI小 = value;
            }
        }

        [Bindable(true)]
        public double? P64折算产量
        {
            get
            {
                return this.currentData.p64折算产量;
            }
            set
            {
                this.currentData.p64折算产量 = value;
            }
        }

        [Bindable(true)]
        public double? P生成标志
        {
            get
            {
                return this.currentData.p生成标志;
            }
            set
            {
                this.currentData.p生成标志 = value;
            }
        }
        #endregion

        public 技术日报内容项(string danwei, string xiangmu)
        {
            this.P01单位 = danwei;
            this.P02项目 = xiangmu;
        }
        protected 技术日报内容项()
        { 
        }
        public void Convert0ToNull()
        {
            if (this.P03合格铁 == 0)
                this.P03合格铁 = null;
            if (this.P04炼钢铁 == 0)
                this.P04炼钢铁= null;
            if (this.P05铸造铁 == 0)
                this.P05铸造铁 = null;
            if (this.P06号外铁 == 0)
                this.P06号外铁  = null;
            if (this.P07合格率 == 0)
                this.P07合格率 = null;
            if (this.P08高炉利用系数 == 0)
                this.P08高炉利用系数 = null;
            if (this.P09高炉实物系数 == 0)
                this.P09高炉实物系数 = null;
            if (this.P10原料矿合计总耗 == 0)
                this.P10原料矿合计总耗 = null;
            if (this.P11原料矿合计单耗 == 0)
                this.P11原料矿合计单耗 = null;
            if (this.P12原料矿机烧 == 0)
                this.P12原料矿机烧 = null;
            if (this.P13原料矿竖炉球 == 0)
                this.P13原料矿竖炉球 = null;
            if (this.P15原料矿其它熟料 == 0)
                this.P15原料矿其它熟料 = null;
            if (this.P16原料矿本溪矿 == 0)
                this.P16原料矿本溪矿 = null;
            if (this.P17原料矿其它生料 == 0)
                this.P17原料矿其它生料 = null;
            if (this.P18废铁总耗 == 0)
                this.P18废铁总耗 = null;
            if (this.P19废铁单耗 == 0)
                this.P19废铁单耗 = null;
            if (this.P20回收率 == 0)
                this.P20回收率 = null;
            if (this.P21熟料比 == 0)
                this.P21熟料比 = null;
            if (this.P22平均风温 == 0)
                this.P22平均风温 = null;
            if (this.P23炉顶温度 == 0)
                this.P23炉顶温度 = null;
            if (this.P24热风压力 == 0)
                this.P24热风压力 = null;
            if (this.P25炉顶压力 == 0)
                this.P25炉顶压力 = null;
            if (this.P26富氧率 == 0)
                this.P26富氧率 = null;
            if (this.P27入炉焦炭总耗 == 0)
                this.P27入炉焦炭总耗 = null;
            if (this.P28入炉焦炭单耗 == 0)
                this.P28入炉焦炭单耗 = null;
            if (this.P29煤粉总耗 == 0)
                this.P29煤粉总耗 = null;
            if (this.P30煤粉单耗 == 0)
                this.P30煤粉单耗 = null;
            if (this.P31焦丁总耗 == 0)
                this.P31焦丁总耗 = null;
            if (this.P32焦丁单耗 == 0)
                this.P32焦丁单耗 = null;
            if (this.P33综合焦炭总耗 == 0)
                this.P33综合焦炭总耗 = null;
            if (this.P34综合焦炭单耗 == 0)
                this.P34综合焦炭单耗 = null;
            if (this.P35综合折算焦比 == 0)
                this.P35综合折算焦比 = null;
            if (this.P36冶炼强度 == 0)
                this.P36冶炼强度 = null;
            if (this.P37焦炭负荷 == 0)
                this.P37焦炭负荷 = null;
            if (this.P38干毛焦 == 0)
                this.P38干毛焦 = null;
            if (this.P39炼钢铁SI == 0)
                this.P39炼钢铁SI = null;
            if (this.P40炼钢铁MN == 0)
                this.P40炼钢铁MN = null;
            if (this.P41炼钢铁P == 0)
                this.P41炼钢铁P = null;
            if (this.P42炼钢铁S == 0)
                this.P42炼钢铁S = null;
            if (this.P43铸造铁SI == 0)
                this.P43铸造铁SI = null;
            if (this.P44铸造铁MN == 0)
                this.P44铸造铁MN = null;
            if (this.P45铸造铁P == 0)
                this.P45铸造铁P = null;
            if (this.P46铸造铁S == 0)
                this.P46铸造铁S = null;
            if (this.P47炉渣碱度 == 0)
                this.P47炉渣碱度 = null;
            if (this.P48休风情况 == 0)
                this.P48休风情况 = null;
            if (this.P49慢风 == 0)
                this.P49慢风 = null;
            if (this.P50坐料次数 == 0)
                this.P50坐料次数 = null;
            if (this.P51悬料次数 == 0)
                this.P51悬料次数 = null;
            if (this.P52崩料次数 == 0)
                this.P52崩料次数 = null;
            if (this.P53风口损坏数大 == 0)
                this.P53风口损坏数大 = null;
            if (this.P54风口损坏数中 == 0)
                this.P54风口损坏数中 = null;
            if (this.P55风口损坏数小 == 0)
                this.P55风口损坏数小 = null;
            if (this.P56渣口损坏数大 == 0)
                this.P56渣口损坏数大 = null;
            if (this.P57渣口损坏数中 == 0)
                this.P57渣口损坏数中 = null;
            if (this.P58渣口损坏数小 == 0)
                this.P58渣口损坏数小 = null;
            if (this.P59本厂铸造SI大 == 0)
                this.P59本厂铸造SI大 = null;
            if (this.P60本厂铸造SI小 == 0)
                this.P60本厂铸造SI小 = null;
            if (this.P61送炼钢厂SI大 == 0)
                this.P61送炼钢厂SI大 = null;
            if (this.P62送炼钢厂SI中 == 0)
                this.P62送炼钢厂SI中 = null;
            if (this.P63送炼钢厂SI小 == 0)
                this.P63送炼钢厂SI小 = null;
                }
    }
}
