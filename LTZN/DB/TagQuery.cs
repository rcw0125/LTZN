using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;
using LTZN;

namespace ZHCDB
{
    public class TagQuery 
    {
          protected OracleConnection conn = new OracleConnection();

          public TagQuery()
        {
            conn.ConnectionString = LTZN.Properties.Settings.Default.ltznConnectionString;
        }
        /// <summary>
        /// 查询班产量
        /// </summary>
        /// <param name="gaolu"></param>
        /// <param name="rq"></param>
        /// <returns></returns>
        public Dictionary<string, double> queryBanChanLiang(int gaolu, DateTime rq)
        {
            Dictionary<string, double> result = new Dictionary<string, double>();
            conn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;

            //各班产量
            cmd.CommandText = "select banci,sum(feliang) from ddluci where  gaolu=:gaolu and trunc(zdsj)=:rq and dksj is not null group by banci";
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            cmd.Parameters.Add(":rq", OracleType.DateTime).Value = rq;
            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0) && !dr.IsDBNull(1))
                {
                    string banci = dr.GetString(0);
                    string tagName = string.Format("{0}产量", banci);
                    result.Add(tagName, dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();

            return result;

        }
    }
}
