using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;

namespace LTZN.技术月报
{
    [DataObjectAttribute()]
    public class 技术月报内容项
    {
        #region 88个属性

        private 技术月报内容项数据 currentData;

        /// <remarks>计量单位－吨</remarks>
        [Bindable(true)]
        public double? P03全铁产量
        {
            get
            {
                return this.currentData.p03全铁产量;
            }
            set
            {  
                this.currentData.p03全铁产量 = value;
            }
        }

        /// <remarks>计量单位－吨</remarks>
        [Bindable(true)]
        public double? P04合格铁
        {
            get
            {
                return this.currentData.p04合格铁;
            }
            set
            {  this.currentData.p04合格铁 = value;
            }
        }

        /// <remarks>计量单位－吨</remarks>
        [Bindable(true)]
        public double? P05制钢铁
        {
            get
            {
                return this.currentData.p05制钢铁;
            }
            set
            {  this.currentData.p05制钢铁 = value;
            }
        }

        /// <remarks>计量单位－吨</remarks>
        [Bindable(true)]
        public double? P06铸造铁
        {
            get
            {
                return this.currentData.p06铸造铁;
            }
            set
            {  this.currentData.p06铸造铁 = value;
            }
        }

        [Bindable(true)]
        public string P01单位
        {
            get
            {
                return this.currentData.p01单位;
            }
            set
            {  this.currentData.p01单位 = value;
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
            {  this.currentData.p02项目 = value;
            }
        }

        /// <remarks>计量单位－吨</remarks>
        [Bindable(true)]
        public double? P07折合产量
        {
            get
            {
                return this.currentData.p07折合产量;
            }
            set
            {  this.currentData.p07折合产量 = value;
            }
        }

        /// <remarks>百分比 ％</remarks>
        [Bindable(true)]
        public double? P08合格率
        {
            get
            {
                return this.currentData.p08合格率;
            }
            set
            {
                this.currentData.p08合格率 = value;
            }
        }

        /// <remarks>百分比 ％</remarks>
        [Bindable(true)]
        public double? P09一级品率
        {
            get
            {
                return this.currentData.p09一级品率;
            }
            set
            {  this.currentData.p09一级品率 = value;
            }
        }

        /// <remarks>百分比 ％</remarks>
        [Bindable(true)]
        public double? P10送炼钢优质铁水率
        {
            get
            {
                return this.currentData.p10送炼钢优质铁水率;
            }
            set
            {  this.currentData.p10送炼钢优质铁水率 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P11高炉利用系数
        {
            get
            {
                return this.currentData.p11高炉利用系数;
            }
            set
            {  this.currentData.p11高炉利用系数 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P12入炉焦炭冶炼强度
        {
            get
            {
                return this.currentData.p12入炉焦炭冶炼强度;
            }
            set
            {  this.currentData.p12入炉焦炭冶炼强度 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P13综合焦炭冶炼强度
        {
            get
            {
                return this.currentData.p13综合焦炭冶炼强度;
            }
            set
            {  this.currentData.p13综合焦炭冶炼强度 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P14折算综合焦比
        {
            get
            {
                return this.currentData.p14折算综合焦比;
            }
            set
            {  this.currentData.p14折算综合焦比 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P16入炉矿品位
        {
            get
            {
                return this.currentData.p16入炉矿品位;
            }
            set
            {  this.currentData.p16入炉矿品位 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P17熟料比
        {
            get
            {
                return this.currentData.p17熟料比;
            }
            set
            {  this.currentData.p17熟料比 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P19综合焦炭负荷
        {
            get
            {
                return this.currentData.p19综合焦炭负荷;
            }
            set
            {  this.currentData.p19综合焦炭负荷 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P18入炉焦炭负荷
        {
            get
            {
                return this.currentData.p18入炉焦炭负荷;
            }
            set
            {  this.currentData.p18入炉焦炭负荷 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P20休风时间
        {
            get
            {
                return this.currentData.p20休风时间;
            }
            set
            {  this.currentData.p20休风时间 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P21休风率
        {
            get
            {
                return this.currentData.p21休风率;
            }
            set
            {  this.currentData.p21休风率 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P22慢风时间
        {
            get
            {
                return this.currentData.p22慢风时间;
            }
            set
            {  this.currentData.p22慢风时间 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P23慢风率
        {
            get
            {
                return this.currentData.p23慢风率;
            }
            set
            {  this.currentData.p23慢风率 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P24煤气成分CO2
        {
            get
            {
                return this.currentData.p24煤气成分CO2;
            }
            set
            {  this.currentData.p24煤气成分CO2 = value;
            }
        }

        /// <remarks>计算公式：CO2/(CO2+CO)</remarks>
        [Bindable(true)]
        public double? P25煤气成分计算1
        {
            get
            {
                return this.currentData.p25煤气成分计算1;
            }
            set
            {  this.currentData.p25煤气成分计算1 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P26生铁表面质量
        {
            get
            {
                return this.currentData.p26生铁表面质量;
            }
            set
            {  this.currentData.p26生铁表面质量 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P27深料线作业率
        {
            get
            {
                return this.currentData.p27深料线作业率;
            }
            set
            {  this.currentData.p27深料线作业率 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P28回收率
        {
            get
            {
                return this.currentData.p28回收率;
            }
            set
            {  this.currentData.p28回收率 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P15折合入炉焦比
        {
            get
            {
                return this.currentData.p15折合入炉焦比;
            }
            set
            {  this.currentData.p15折合入炉焦比 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P29冷风流量
        {
            get
            {
                return this.currentData.p29冷风流量;
            }
            set
            {  this.currentData.p29冷风流量 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P30平均风温
        {
            get
            {
                return this.currentData.p30平均风温;
            }
            set
            {  this.currentData.p30平均风温 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P40风口大套损坏数
        {
            get
            {
                return this.currentData.p40风口大套损坏数;
            }
            set
            {  this.currentData.p40风口大套损坏数 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P31风速
        {
            get
            {
                return this.currentData.p31风速;
            }
            set
            {  this.currentData.p31风速 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P32热风压力
        {
            get
            {
                return this.currentData.p32热风压力;
            }
            set
            {  this.currentData.p32热风压力 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P33炉顶压力
        {
            get
            {
                return this.currentData.p33炉顶压力;
            }
            set
            {  this.currentData.p33炉顶压力 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P34炉顶温度
        {
            get
            {
                return this.currentData.p34炉顶温度;
            }
            set
            {  this.currentData.p34炉顶温度 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P35富氧率
        {
            get
            {
                return this.currentData.p35富氧率;
            }
            set
            {  this.currentData.p35富氧率 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P36高压率
        {
            get
            {
                return this.currentData.p36高压率;
            }
            set
            {  this.currentData.p36高压率 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P37悬料次数
        {
            get
            {
                return this.currentData.p37悬料次数;
            }
            set
            {  this.currentData.p37悬料次数 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P38坐料次数
        {
            get
            {
                return this.currentData.p38坐料次数;
            }
            set
            {  this.currentData.p38坐料次数 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P39崩料次数
        {
            get
            {
                return this.currentData.p39崩料次数;
            }
            set
            {  this.currentData.p39崩料次数 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P41风口中套损坏数
        {
            get
            {
                return this.currentData.p41风口中套损坏数;
            }
            set
            {  this.currentData.p41风口中套损坏数 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P51竖炉球消耗
        {
            get
            {
                return this.currentData.p51竖炉球消耗;
            }
            set
            {  this.currentData.p51竖炉球消耗 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P42风口小套损坏数
        {
            get
            {
                return this.currentData.p42风口小套损坏数;
            }
            set
            {  this.currentData.p42风口小套损坏数 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P43渣口大套损坏数
        {
            get
            {
                return this.currentData.p43渣口大套损坏数;
            }
            set
            {  this.currentData.p43渣口大套损坏数 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P44渣口中套损坏数
        {
            get
            {
                return this.currentData.p44渣口中套损坏数;
            }
            set
            {  this.currentData.p44渣口中套损坏数 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P45渣口小套损坏数
        {
            get
            {
                return this.currentData.p45渣口小套损坏数;
            }
            set
            {  this.currentData.p45渣口小套损坏数 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P46渣铁比
        {
            get
            {
                return this.currentData.p46渣铁比;
            }
            set
            {  this.currentData.p46渣铁比 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P47灰铁比
        {
            get
            {
                return this.currentData.p47灰铁比;
            }
            set
            {  this.currentData.p47灰铁比 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P48原料总耗
        {
            get
            {
                return this.currentData.p48原料总耗;
            }
            set
            {  this.currentData.p48原料总耗 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P50机烧消耗
        {
            get
            {
                return this.currentData.p50机烧消耗;
            }
            set
            {  this.currentData.p50机烧消耗 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P52印度球消耗
        {
            get
            {
                return this.currentData.p52印度球消耗;
            }
            set
            {  this.currentData.p52印度球消耗 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P53其它熟料消耗
        {
            get
            {
                return this.currentData.p53其它熟料消耗;
            }
            set
            {  this.currentData.p53其它熟料消耗 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P54本溪矿消耗
        {
            get
            {
                return this.currentData.p54本溪矿消耗;
            }
            set
            {  this.currentData.p54本溪矿消耗 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P55其它生料消耗
        {
            get
            {
                return this.currentData.p55其它生料消耗;
            }
            set
            {  this.currentData.p55其它生料消耗 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P56废铁总耗
        {
            get
            {
                return this.currentData.p56废铁总耗;
            }
            set
            {  this.currentData.p56废铁总耗 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P57废铁单耗
        {
            get
            {
                return this.currentData.p57废铁单耗;
            }
            set
            {  this.currentData.p57废铁单耗 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P49原料单耗
        {
            get
            {
                return this.currentData.p49原料单耗;
            }
            set
            {  this.currentData.p49原料单耗 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P58金属收得率
        {
            get
            {
                return this.currentData.p58金属收得率;
            }
            set
            {
                this.currentData.p58金属收得率 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P59综合焦炭总耗
        {
            get
            {
                return this.currentData.p59综合焦炭总耗;
            }
            set
            {
                this.currentData.p59综合焦炭总耗 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P68燃料比
        {
            get
            {
                return this.currentData.p68燃料比;
            }
            set
            {
                this.currentData.p68燃料比 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P67焦丁单耗
        {
            get
            {
                return this.currentData.p67焦丁单耗;
            }
            set
            {
                this.currentData.p67焦丁单耗 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P66焦丁总耗
        {
            get
            {
                return this.currentData.p66焦丁总耗;
            }
            set
            {
                this.currentData.p66焦丁总耗 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P64入炉焦炭单耗
        {
            get
            {
                return this.currentData.p64入炉焦炭单耗;
            }
            set
            {
                this.currentData.p64入炉焦炭单耗 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P65煤粉总耗
        {
            get
            {
                return this.currentData.p65煤粉总耗;
            }
            set
            {
                this.currentData.p65煤粉总耗 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P63入炉焦炭总耗
        {
            get
            {
                return this.currentData.p63入炉焦炭总耗;
            }
            set
            {
                this.currentData.p63入炉焦炭总耗 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P62干焦粉
        {
            get
            {
                return this.currentData.p62干焦粉;
            }
            set
            {
                this.currentData.p62干焦粉 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P61干毛焦
        {
            get
            {
                return this.currentData.p61干毛焦;
            }
            set
            {
                this.currentData.p61干毛焦 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P60七号称
        {
            get
            {
                return this.currentData.p60七号称;
            }
            set
            {
                this.currentData.p60七号称 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P69铁成分SI
        {
            get
            {
                return this.currentData.p69铁成分Si;
            }
            set
            {
                this.currentData.p69铁成分Si = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P70铁成分MN
        {
            get
            {
                return this.currentData.p70铁成分Mn;
            }
            set
            {
                this.currentData.p70铁成分Mn = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P71铁成分P
        {
            get
            {
                return this.currentData.p71铁成分P;
            }
            set
            {
                this.currentData.p71铁成分P = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P72铁成分S
        {
            get
            {
                return this.currentData.p72铁成分S;
            }
            set
            {
                this.currentData.p72铁成分S = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P73渣成分CAO
        {
            get
            {
                return this.currentData.p73渣成分CaO;
            }
            set
            {
                this.currentData.p73渣成分CaO = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P74渣成分SIO2
        {
            get
            {
                return this.currentData.p74渣成分SIO2;
            }
            set
            {
                this.currentData.p74渣成分SIO2 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P75渣成分AL2O3
        {
            get
            {
                return this.currentData.p75渣成分Al2O3;
            }
            set
            {
                this.currentData.p75渣成分Al2O3 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P76渣成分MGO
        {
            get
            {
                return this.currentData.p76渣成分MgO;
            }
            set
            {
                this.currentData.p76渣成分MgO = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P77渣成分FEO
        {
            get
            {
                return this.currentData.p77渣成分FeO;
            }
            set
            {
                this.currentData.p77渣成分FeO = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P78渣成分S
        {
            get
            {
                return this.currentData.p78渣成分S;
            }
            set
            {
                this.currentData.p78渣成分S = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P79含SI偏差CAO
        {
            get
            {
                return this.currentData.p79含Si偏差CaO;
            }
            set
            {
                this.currentData.p79含Si偏差CaO = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P80含SI偏差制铁量
        {
            get
            {
                return this.currentData.p80含Si偏差制铁量;
            }
            set
            {
                this.currentData.p80含Si偏差制铁量 = value;
            }
        }

        /// <remarks></remarks>
        [Bindable(true)]
        public double? P81含SI偏差铸造铁
        {
            get
            {
                return this.currentData.p81含Si偏差铸造铁;
            }
            set
            {
                this.currentData.p81含Si偏差铸造铁 = value;
            }
        }

        [Bindable(true)]
        public double? 综合焦炭单耗
        {
            get
            {
                return this.currentData.综合焦炭单耗;
            }
            set
            {
                this.currentData.综合焦炭单耗 = value;
            }
        }

        [Bindable(true)]
        public double? 煤粉单耗
        {
            get
            {
                return this.currentData.煤粉单耗;
            }
            set
            {
                this.currentData.煤粉单耗 = value;
            }
        }

        [Bindable(true)]
        public double? 一级铁
        {
            get
            {
                return this.currentData.一级铁;
            }
            set
            {
                this.currentData.一级铁 = value;
            }
        }

        [Bindable(true)]
        public double? 优质铁
        {
            get
            {
                return this.currentData.优质铁;
            }
            set
            {
                this.currentData.优质铁 = value;
            }
        }

        [Bindable(true)]
        public double? 瓦斯灰
        {
            get
            {
                return this.currentData.瓦斯灰;
            }
            set
            {
                this.currentData.瓦斯灰 = value;
            }
        }

        [Bindable(true)]
        public double? 球团矿
        {
            get
            {
                return this.currentData.球团矿;
            }
            set
            {
                this.currentData.球团矿 = value;
            }
        }

        [Bindable(true)]
        public double? 国内球团矿
        {
            get
            {
                return this.currentData.国内球团矿;
            }
            set
            {
                this.currentData.国内球团矿 = value;
            }
        }

        [Bindable(true)]
        public double? 进口球团矿
        {
            get
            {
                return this.currentData.进口球团矿;
            }
            set
            {
                this.currentData.进口球团矿 = value;
            }
        }

        [Bindable(true)]
        public double? 其它块矿
        {
            get
            {
                return this.currentData.其它块矿;
            }
            set
            {
                this.currentData.其它块矿 = value;
            }
        }

        [Bindable(true)]
        public double? 高钛球团矿
        {
            get
            {
                return this.currentData.高钛球团矿;
            }
            set
            {
                this.currentData.高钛球团矿 = value;
            }
        }

        [Bindable(true)]
        public double? 高品位钛球
        {
            get
            {
                return this.currentData.高品位钛球;
            }
            set
            {
                this.currentData.高品位钛球 = value;
            }
        }

        #endregion

        public 技术月报内容项(string 单位,string 项目)
        {
            this.currentData.p01单位 = 单位;
            this.currentData.p02项目 = 项目;
        }
        protected 技术月报内容项()
        { 
        }
        public void Convert0ToNull()
        {
            if (this.P03全铁产量  == 0)
                this.P03全铁产量 = null;
            if (this.P04合格铁 == 0)
                this.P04合格铁= null;
            if (this.P05制钢铁 == 0)
                this.P05制钢铁 = null;
            if (this.P06铸造铁 == 0)
                this.P06铸造铁  = null;
            if (this.P07折合产量 == 0)
                this.P07折合产量 = null;
            if (this.P08合格率 == 0)
                this.P08合格率 = null;
            if (this.P09一级品率 == 0)
                this.P09一级品率 = null;
            if (this.P10送炼钢优质铁水率 == 0)
                this.P10送炼钢优质铁水率 = null;
            if (this.P11高炉利用系数 == 0)
                this.P11高炉利用系数 = null;
            if (this.P12入炉焦炭冶炼强度 == 0)
                this.P12入炉焦炭冶炼强度 = null;
            if (this.P13综合焦炭冶炼强度 == 0)
                this.P13综合焦炭冶炼强度 = null;
            if (this.P14折算综合焦比 == 0)
                this.P14折算综合焦比 = null;
            if (this.P15折合入炉焦比 == 0)
                this.P15折合入炉焦比 = null;
            if (this.P16入炉矿品位 == 0)
                this.P16入炉矿品位 = null;
            if (this.P17熟料比 == 0)
                this.P17熟料比 = null;
            if (this.P18入炉焦炭负荷 == 0)
                this.P18入炉焦炭负荷 = null;
            if (this.P19综合焦炭负荷 == 0)
                this.P19综合焦炭负荷 = null;
            if (this.P20休风时间 == 0)
                this.P20休风时间 = null;
            if (this.P21休风率 == 0)
                this.P21休风率 = null;
            if (this.P22慢风时间 == 0)
                this.P22慢风时间 = null;
            if (this.P23慢风率 == 0)
                this.P23慢风率 = null;
            if (this.P24煤气成分CO2 == 0)
                this.P24煤气成分CO2 = null;
            if (this.P25煤气成分计算1 == 0)
                this.P25煤气成分计算1 = null;
            if (this.P27深料线作业率 == 0)
                this.P27深料线作业率 = null;
            if (this.P28回收率 == 0)
                this.P28回收率 = null;
            if (this.P29冷风流量 == 0)
                this.P29冷风流量 = null;
            if (this.P30平均风温 == 0)
                this.P30平均风温 = null;
            if (this.P31风速 == 0)
                this.P31风速 = null;
            if (this.P32热风压力  == 0)
                this.P32热风压力 = null;
            if (this.P33炉顶压力 == 0)
                this.P33炉顶压力 = null;
            if (this.P34炉顶温度 == 0)
                this.P34炉顶温度 = null;
            if (this.P35富氧率 == 0)
                this.P35富氧率 = null;
            if (this.P36高压率 == 0)
                this.P36高压率 = null;
            if (this.P37悬料次数 == 0)
                this.P37悬料次数 = null;
            if (this.P38坐料次数 == 0)
                this.P38坐料次数 = null;
            if (this.P39崩料次数 == 0)
                this.P39崩料次数 = null;
            if (this.P40风口大套损坏数 == 0)
                this.P40风口大套损坏数 = null;
            if (this.P41风口中套损坏数 == 0)
                this.P41风口中套损坏数 = null;
            if (this.P42风口小套损坏数 == 0)
                this.P42风口小套损坏数 = null;
            if (this.P43渣口大套损坏数 == 0)
                this.P43渣口大套损坏数 = null;
            if (this.P44渣口中套损坏数 == 0)
                this.P44渣口中套损坏数 = null;
            if (this.P45渣口小套损坏数 == 0)
                this.P45渣口小套损坏数 = null;
            if (this.P46渣铁比 == 0)
                this.P46渣铁比 = null;
            if (this.P47灰铁比 == 0)
                this.P47灰铁比 = null;
            if (this.P48原料总耗 == 0)
                this.P48原料总耗 = null;
            if (this.P49原料单耗 == 0)
                this.P49原料单耗 = null;
            if (this.P50机烧消耗 == 0)
                this.P50机烧消耗 = null;
            if (this.P51竖炉球消耗 == 0)
                this.P51竖炉球消耗 = null;
            if (this.P52印度球消耗 == 0)
                this.P52印度球消耗 = null;
            if (this.P53其它熟料消耗 == 0)
                this.P53其它熟料消耗 = null;
            if (this.P54本溪矿消耗 == 0)
                this.P54本溪矿消耗 = null;
            if (this.P55其它生料消耗 == 0)
                this.P55其它生料消耗 = null;
            if (this.P56废铁总耗 == 0)
                this.P56废铁总耗 = null;
            if (this.P57废铁单耗 == 0)
                this.P57废铁单耗 = null;
            if (this.P58金属收得率 == 0)
                this.P58金属收得率 = null;
            if (this.P59综合焦炭总耗 == 0)
                this.P59综合焦炭总耗 = null;
            if (this.P60七号称 == 0)
                this.P60七号称 = null;
            if (this.P61干毛焦 == 0)
                this.P61干毛焦 = null;
            if (this.P62干焦粉 == 0)
                this.P62干焦粉 = null;
            if (this.P63入炉焦炭总耗 == 0)
                this.P63入炉焦炭总耗 = null;
            if (this.P64入炉焦炭单耗 == 0)
                this.P64入炉焦炭单耗 = null;
            if (this.P65煤粉总耗 == 0)
                this.P65煤粉总耗 = null;
            if (this.P66焦丁总耗 == 0)
                this.P66焦丁总耗 = null;
            if (this.P67焦丁单耗 == 0)
                this.P67焦丁单耗 = null;
            if (this.P68燃料比 == 0)
                this.P68燃料比 = null;
            if (this.P69铁成分SI == 0)
                this.P69铁成分SI = null;
            if (this.P70铁成分MN == 0)
                this.P70铁成分MN = null;
            if (this.P71铁成分P == 0)
                this.P71铁成分P = null;
            if (this.P72铁成分S == 0)
                this.P72铁成分S = null;
            if (this.P73渣成分CAO == 0)
                this.P73渣成分CAO = null;
            if (this.P74渣成分SIO2 == 0)
                this.P74渣成分SIO2 = null;
            if (this.P75渣成分AL2O3 == 0)
                this.P75渣成分AL2O3 = null;
            if (this.P76渣成分MGO == 0)
                this.P76渣成分MGO = null;
            if (this.P77渣成分FEO== 0)
                this.P77渣成分FEO = null;
            if (this.P78渣成分S == 0)
                this.P78渣成分S = null;
            if (this.P79含SI偏差CAO == 0)
                this.P79含SI偏差CAO = null;
            if (this.P80含SI偏差制铁量 == 0)
                this.P80含SI偏差制铁量 = null;
            if (this.P81含SI偏差铸造铁 == 0)
                this.P81含SI偏差铸造铁 = null;
            if (this.综合焦炭单耗 == 0)
                this.综合焦炭单耗 = null;
            if (this.煤粉单耗 == 0)
                this.煤粉单耗 = null;

              if (this.球团矿 == 0)
                this.球团矿 = null;

              if (this.国内球团矿 == 0)
                  this.国内球团矿 = null;

              if (this.进口球团矿 == 0)
                  this.进口球团矿 = null;

              if (this.其它块矿 == 0)
                  this.其它块矿 = null;

              if (this.高钛球团矿 == 0)
                  this.高钛球团矿 = null;

              if (this.高品位钛球 == 0)
                  this.高品位钛球 = null;

              if (this.其它熔剂 == 0)
                  this.其它熔剂 = null;
                

                }

          

    }
    }

