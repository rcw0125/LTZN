using System;
namespace LTZN.技术月报
{


    partial class 技术月报表
    {
        partial class 技术月报DataTable
        {
            public 技术月报Row Add技术月报Row(
                        int P年,
                        int P月,
                        string P01单位,
                        string P02项目)
            {
                技术月报Row row技术月报Row = ((技术月报Row)(this.NewRow()));
                row技术月报Row.ItemArray = new object[] {
                        P年,
                        P月,
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
                this.Rows.Add(row技术月报Row);
                return row技术月报Row;
            }
            public void 计算()
            {
                foreach (技术月报Row r in this.Rows)
                {
                    r.计算();
                }
            }

        }
        public partial class 技术月报Row
        {
            public void 计算()
            {
                if (收入含铁 > 0)
                    P58金属收得率 = 支出含铁 * 100 / 收入含铁;
                P58金属收得率 = decimal.Parse(P58金属收得率.ToString("N2"));
                if (P03全铁产量 > 0)
                    P09一级品率 = 一级铁 * 100 / P03全铁产量;
                P09一级品率 = decimal.Parse(P09一级品率.ToString("N2"));
            }
        }
    }
}
