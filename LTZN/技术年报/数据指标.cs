using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.技术年报
{
    class 数据指标
    {
        private double f全铁产量;

        private double f合格铁;

        private double f炼钢铁;

        private double f折算产量;

        private double f一级铁;

        private double f炼钢铁一级铁;

        private double f原料矿;

        private double f人造矿;

        private double f碎铁;

        private double f综合焦炭;

        private double f煤粉;

        private double f入炉焦炭;


        public double 全铁产量
        {
            get
            {
                return f全铁产量;
            }
        }

        public double 合格铁
        {
            get
            {
                return f合格铁;
            }
        }

        public double 出格铁
        {
            get
            {
               return f全铁产量 - f合格铁;
            }
        }

        public double 炼钢铁
        {
            get
            {
               return f炼钢铁;
            }
        }

        public double 铸造铁
        {
            get
            {
               return f全铁产量 - f炼钢铁;
            }
        }

        public double 折算产量
        {
            get
            {
               return f折算产量;
            }
        }

        public double 生铁合格率
        {
            get
            {
               return double.Parse(((double)(f合格铁 / f全铁产量)).ToString("N2"));
            }
        }

        public double 生铁一级品率
        {
            get
            {
                return double.Parse(((double)(f一级铁 / f全铁产量)).ToString("N2"));
            }
        }

        public double 炼钢铁一级品率
        {
            get
            {
                return double.Parse(((double)(f炼钢铁一级铁 / f全铁产量)).ToString("N2"));
            }
        }

        public double 铸造铁一级品率
        {
            get
            {
                return double.Parse(((double)((f一级铁 - f炼钢铁一级铁) / f全铁产量)).ToString("N2"));
            }
        }

        public double 原料矿石
        {
            get
            {
                return double.Parse(((double)(f原料矿*1000 / f全铁产量)).ToString("N0"));
            }
        }

        public double 人造矿石
        {
            get
            {
                return double.Parse(((double)(f人造矿 * 1000 / f全铁产量)).ToString("N0"));
            }
        }

        public double 天然矿石
        {
            get
            {
                return double.Parse(((double)((f原料矿-f人造矿) * 1000 / f全铁产量)).ToString("N0"));
            }
        }

        public double 碎铁
        {
            get
            {
                return double.Parse(((double)(f碎铁 * 1000 / f全铁产量)).ToString("N0"));
            }
        }

        public double 综合焦比
        {
            get
            {
                return double.Parse(((double)(f综合焦炭 * 1000 / f全铁产量)).ToString("N0"));
            }
        }

        public double 煤粉比
        {
            get
            {
                return double.Parse(((double)(f煤粉 * 1000 / f全铁产量)).ToString("N0"));
            }
        }

        public double 入炉焦比
        {
            get
            {
                return double.Parse(((double)(f入炉焦炭 * 1000 / f全铁产量)).ToString("N0"));
            }
        }

        public double 折算综合焦比
        {
            get
            {
                return double.Parse(((double)(f综合焦炭 * 1000 / f折算产量)).ToString("N0"));
            }
        }

        public double 折算入炉焦比
        {
            get
            {
                return double.Parse(((double)(f入炉焦炭 * 1000 / f折算产量)).ToString("N0"));
            }
        }




    }
}
