namespace LTZN.全厂日消耗
{


    partial class 全厂日消耗数据集
    {

        partial class 全厂日消耗DataTable
        {
            public 全厂日消耗Row Add全厂日消耗Row(System.DateTime P01日期)
            {
                全厂日消耗Row row全厂日消耗Row = ((全厂日消耗Row)(this.NewRow()));
                row全厂日消耗Row.ItemArray = new object[] {
                        P01日期,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull};
                this.Rows.Add(row全厂日消耗Row);
                return row全厂日消耗Row;
            }
        }

        partial class 日原料成分DataTable
        {
            public 日原料成分Row Add日原料成分Row(System.DateTime P01日期, string P02名称)
            {
                日原料成分Row row日原料成分Row = ((日原料成分Row)(this.NewRow()));
                row日原料成分Row.ItemArray = new object[] {
                        P01日期,
                        P02名称,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull,
                        System.Convert.DBNull};
                this.Rows.Add(row日原料成分Row);
                return row日原料成分Row;
            }
        }
    }
}
