namespace LTZN.原料消耗
{


    partial class 原料消耗数据集
    {
        partial class 原料消耗DataTable
        {
            public 原料消耗Row Add原料消耗Row(System.DateTime 日期, decimal 高炉)
            {
                原料消耗Row row原料消耗Row = ((原料消耗Row)(this.NewRow()));
                row原料消耗Row.ItemArray = new object[] {
                        日期,
                        高炉,
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
                        "其它熟料",
                        System.Convert.DBNull,
                        "其它生料"};
                this.Rows.Add(row原料消耗Row);
                return row原料消耗Row;
            }
        }
    }
}

namespace LTZN.原料消耗.原料消耗数据集TableAdapters
{


    public partial class 原料消耗TableAdapter
    {
    }
}
