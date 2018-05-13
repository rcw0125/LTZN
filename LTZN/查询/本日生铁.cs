using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;
using LTZN.CalFramework;
using LTZN.Repository;

namespace LTZN.查询
{
    public class 本日生铁 : QueryModel
    {
        
        public 本日生铁():base("本日生铁")
        {
        }

        public void getData(DateTime rq)
        {
            this.myModel.Init();

            conn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;

            cmd.CommandText = "select gaolu,banci,sum(feliang) as liang from ddluci where trunc(zdsj)=:rq and dksj is not null group by gaolu,banci";
            cmd.Parameters.Add(":rq", OracleType.DateTime).Value = rq;
            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0) && !dr.IsDBNull(1) && !dr.IsDBNull(2))//&& !dr.IsDBNull(3) && !dr.IsDBNull(4) && !dr.IsDBNull(5) && !dr.IsDBNull(6) && !dr.IsDBNull(6) && !dr.IsDBNull(7) && !dr.IsDBNull(8))
                {
                    int gaolu = dr.GetInt32(0);
                    string banci = dr.GetString(1);
                    this.SetValue("高炉" + gaolu.ToString() + "." + banci, dr.GetDouble(2));
                  
                }
            }
            dr.Close();
            //2012-6-12 BY lhl添加 新增FE中FE

            cmd.CommandText = "select gaolu,sum(feliang) as liang, sum(feliang*fep)/100 as p,sum(feliang*fes)/100 as s,sum(feliang*fec)/100 as c,sum(feliang*fesi)/100 as si,sum(feliang*femn)/100 as mn,sum(feliang*feti)/100  from ddluci where trunc(zdsj)=:rq and dksj is not null group by gaolu";
            cmd.Parameters.Add(":rq", OracleType.DateTime).Value = rq;
           dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0) && !dr.IsDBNull(1))//&& !dr.IsDBNull(3) && !dr.IsDBNull(4) && !dr.IsDBNull(5) && !dr.IsDBNull(6) && !dr.IsDBNull(6) && !dr.IsDBNull(7) && !dr.IsDBNull(8))
                {
                    int gaolu = dr.GetInt32(0);
                     double qita = 0;
                    if (!dr.IsDBNull(3))
                    {
                        qita = dr.GetDouble(3) + qita;
                    }
                    if (!dr.IsDBNull(4))
                    {
                        qita = dr.GetDouble(4) + qita;
                    }
                    if (!dr.IsDBNull(5))
                    {
                        qita = dr.GetDouble(5) + qita;
                    }
                    if (!dr.IsDBNull(6))
                    {
                        qita = dr.GetDouble(6) + qita;
                    }
                    if (!dr.IsDBNull(7))
                    {
                        qita = dr.GetDouble(7) + qita;
                    }
                    if (!dr.IsDBNull(2))
                    {
                        qita = dr.GetDouble(2) + qita;
                    }
                    double febili = Convert.ToDouble(((1 - qita / dr.GetDouble(1)) * 100).ToString("00.00"));
                    this.SetValue("高炉" + gaolu.ToString() + "." + "铁中铁", febili);
                }
            }
            dr.Close();
            //Si大小分类产量
            StringBuilder sqlSb = new StringBuilder();
            sqlSb.Append("select gaolu,");
            sqlSb.Append("case when fesi<=0.7  then 'L04' ");
            sqlSb.Append("when  fesi<=0.85 then 'L08' ");
            sqlSb.Append("when  fesi<=1.25 then 'L10' ");
            sqlSb.Append("when  fesi<1.6 then 'Z14' ");
            sqlSb.Append("when  fesi<=2 then 'Z18' ");
            sqlSb.Append("when  fesi<=2.4 then 'Z22' ");
            sqlSb.Append("ELSE 'Z26' ");
            sqlSb.Append("end,");
            sqlSb.Append("sum(feliang) from ddluci where trunc(zdsj)=:rq and dksj is not null group by gaolu,");
            sqlSb.Append("case when fesi<=0.7  then 'L04' ");
            sqlSb.Append("when  fesi<=0.85 then 'L08' ");
            sqlSb.Append("when  fesi<=1.25 then 'L10' ");
            sqlSb.Append("when  fesi<1.6 then 'Z14' ");
            sqlSb.Append("when  fesi<=2 then 'Z18' ");
            sqlSb.Append("when  fesi<=2.4 then 'Z22' ");
            sqlSb.Append("else 'Z26' ");
            sqlSb.Append("end");
            cmd.CommandText = sqlSb.ToString();
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    if (!dr.IsDBNull(1))
                        this.SetValue(string.Format("高炉{0}.{1}", gaolu, dr.GetString(1)), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                }
            }
            dr.Close();
            conn.Close();
        }
    
    }
}
