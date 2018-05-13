using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;

namespace LTZN.高炉燃料比综合分析
{
    class LegendEnviroment
    {

        public static void loadZichanj(DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select trunc(sj),trunc(avg(huifen),2) from DDJT t where trunc(t.sj)>=trunc(:sjBegin) and trunc(t.sj)<=trunc(:sjEnd) and t.mc like '自产%' group by trunc(sj) order by trunc(sj)";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;

            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));
                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }

        public static void loadZichanjShuiFen(DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select trunc(sj),trunc(avg(shuif),2) from DDJT t where trunc(t.sj)>=trunc(:sjBegin) and trunc(t.sj)<=trunc(:sjEnd) and t.mc like '自产%' group by trunc(sj) order by trunc(sj)";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;

            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));
                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }
        public static void loadWaiGouj(DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select trunc(sj),trunc(avg(huifen),2) from DDJT t where trunc(t.sj)>=trunc(:sjBegin) and trunc(t.sj)<=trunc(:sjEnd) and t.mc like '外进%' group by trunc(sj) order by trunc(sj)";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;

            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));
                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }
        public static void loadMeiFenh(DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select trunc(sj),trunc(avg(HUIFEN),2) from DDMF t where trunc(t.sj)>=trunc(:sjBegin) and trunc(t.sj)<=trunc(:sjEnd) and t.pz like '煤粉%' group by trunc(sj) order by trunc(sj)";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;

            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));
                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }
        public static void loadMeiFenhuifa(DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select trunc(sj),trunc(avg(HUIFA),2) from DDMF t where trunc(t.sj)>=trunc(:sjBegin) and trunc(t.sj)<=trunc(:sjEnd) and t.pz like '煤粉%' group by trunc(sj) order by trunc(sj)";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;

            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));
                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }
        public static void loadMeiFenShuifen(DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select trunc(sj),trunc(avg(SF),2) from DDMF t where trunc(t.sj)>=trunc(:sjBegin) and trunc(t.sj)<=trunc(:sjEnd) and t.pz like '煤粉%' group by trunc(sj) order by trunc(sj)";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;

            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));
                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }
        public static void loadWaiGoujtshuifen(DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select trunc(sj),trunc(avg(shuif),2) from DDJT t where trunc(t.sj)>=trunc(:sjBegin) and trunc(t.sj)<=trunc(:sjEnd) and t.mc like '外进%' group by trunc(sj) order by trunc(sj)";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;

            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));
                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }

        public static void loadSJKPinWei(DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select trunc(sj)，trunc(avg(tfe),2) from DDYL t  WHERE  TRUNC(SJ)>=:sjBegin AND TRUNC(SJ)<=:sjEnd and MC LIKE '%大烧' group by trunc(sj) order by trunc(sj)";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;

            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));
                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }
      // 块矿品味
        public static void loadkuaikuangPinWei(DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select trunc(sj)，trunc(avg(tfe),2) from DDYL t  WHERE  TRUNC(SJ)>=:sjBegin AND TRUNC(SJ)<=:sjEnd and MC LIKE '%块矿%' group by trunc(sj) order by trunc(sj)";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;

            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));
                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }
        //五高炉烧结品味
        public static void loadSJKPinWei5(DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select trunc(sj)，trunc(avg(tfe),2) from DDYL t  WHERE  TRUNC(SJ)>=:sjBegin AND TRUNC(SJ)<=:sjEnd and cang like '5#%' and mc like '%大烧%' group by trunc(sj) order by trunc(sj)";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;

            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));
                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }

        //球团矿品味
        public static void loadQTKPinWei(DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select trunc(sj)，trunc(avg(tfe),2) from DDYL t  WHERE  TRUNC(SJ)>=:sjBegin AND TRUNC(SJ)<=:sjEnd and MC LIKE '竖%' group by trunc(sj) order by trunc(sj)";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;

            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));
                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }

      
        //澳矿品味
        public static void loadAOKPinWei(DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select trunc(sj)，trunc(tfe,2) from DDYL t  WHERE  MC LIKE '澳%'";
            OracleCommand cmd = new OracleCommand(sql, conn);
          
            OracleDataReader dr = cmd.ExecuteReader();
            TimeSpan i = t2.Subtract(t1);
            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    for (int a = 0; a < i.Days+1; a++)
                    {
                        dts.Add(t1.AddDays(a));
                        vs.Add(dr.GetDouble(1));
                    }
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }

        //球团品味
        public static void loadQTKKPinWei(DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select trunc(sj)，trunc(avg(tfe),2) from DDYL t  WHERE  TRUNC(SJ)>=:sjBegin AND TRUNC(SJ)<=:sjEnd and MC LIKE '其它%' group by trunc(sj) order by trunc(sj)";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;

            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));
                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }
        //钛球品味
        public static void loadTQPinWei(DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select trunc(sj)，trunc(avg(tfe),2) from DDYL t  WHERE  TRUNC(SJ)>=:sjBegin AND TRUNC(SJ)<=:sjEnd and MC LIKE '钛球%' group by trunc(sj) order by trunc(sj)";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;

            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));
                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }
       
        //load 高炉原料种类
        public static void loadGLYL(string mc,out List<string> s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<string> s1 = new List<string>();
            string sql = "select ylmc from ylfl where ylzl=:ylzl order by ylmc ";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":ylzl", OracleType.NVarChar).Value = mc;
            OracleDataReader dr = cmd.ExecuteReader();
           
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {

                    s1.Add(dr.GetString(0));
                }
            }
            dr.Close();
            conn.Close();
            s = s1;
        }
        // 入炉原料结构配比
        public static void loadRlylPeiBi(int gaolu,string ylmc,DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select a.sj,trunc(a.lj*100/b.zong,2) from " +
                        "(select  (nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0)) as lj,sj from rbxiaohao where TRUNC(SJ)>=:sjBegin and TRUNC(SJ)<=:sjEnd and gaolu=:gaolu and mc like :ylmc) a," +
                        "(select sum(nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0)) as zong,sj from rbxiaohao where TRUNC(SJ)>=:sjBegin and TRUNC(SJ)<=:sjEnd and gaolu=:gaolu " +
                        "and  mc in (select ylmc from ylfl where ylzl in ('含铁原料','熔剂')) group by sj) b where a.sj =b.sj order by sj";

            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            cmd.Parameters.Add(":ylmc", OracleType.NVarChar).Value ="%" +ylmc+"%";
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));
                  
                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }


        class CalYuanLiao
        {
            public string MingCheng
            {
                get;
                set;
            }

            public double Bili
            {
                get;
                set;
            }

            public double Pinwei
            {
                get;
                set;
            }
        }

        class CalYuanLiaoTian : List<CalYuanLiao>
        {
            public DateTime Riqi
            {
                get;
                set;
            }

            public void SetBili(string mc, double bili)
            {
                GetYuanliao(mc).Bili = bili;
            }

            public void SetPinwei(string mc,double pinwei)
            {
                GetYuanliao(mc).Pinwei = pinwei;
            }

            public CalYuanLiao GetYuanliao(string mc)
            {
                foreach (var item in this)
                {
                    if (item.MingCheng == mc)
                    {
                        return item;
                    }
                }
                CalYuanLiao cyl = new CalYuanLiao();
                cyl.MingCheng = mc;
                this.Add(cyl);
                return cyl;
            }

            public double HuiShouLv
            {
                get {
                    double result = 0;
                    foreach (var item in this)
                    {
                        result += item.Bili * item.Pinwei;
                    }
                    result /= 100;
                    return result;
                }
            }
        }


        class CalYuanLiaoTianTable : List<CalYuanLiaoTian>
        {

            public void SetBili(DateTime riqi,string mc, double bili)
            {
               GetCalYuanLiaoTian(riqi).GetYuanliao(mc).Bili = bili;
            }

            public void SetPinwei(DateTime riqi, string mc, double pinwei)
            {
                GetCalYuanLiaoTian(riqi).GetYuanliao(mc).Pinwei = pinwei;
            }

            public CalYuanLiaoTian GetCalYuanLiaoTian(DateTime riqi)
            {
                foreach (var item in this)
                {
                    if (item.Riqi == riqi)
                    {
                        return item;
                    }
                }
                CalYuanLiaoTian cyl = new CalYuanLiaoTian();
                cyl.Riqi = riqi;
                this.Add(cyl);
                return cyl;
            }
        }

        // 入炉计算回收率
        public static void loadLlhsl(int gaolu, DateTime t1, DateTime t2, out DateTime[] sT,out double[] s)
        {

            CalYuanLiaoTianTable hslTable = new CalYuanLiaoTianTable();

            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
          
           
            List<double> Hslvs = new List<double>();

            string sql = "select a.sj,mc,trunc(a.lj*100/b.zong,2) from "+
                "(select (nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0)) as lj,sj,mc from rbxiaohao where TRUNC(SJ)>=:sjBegin and TRUNC(SJ)<=:sjEnd and gaolu=:gaolu and  mc in (select ylmc from ylfl where ylzl in ('含铁原料'))) a," +
                "(select sum(nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0)) as zong ,sj from rbxiaohao where TRUNC(SJ)>=:sjBegin and TRUNC(SJ)<=:sjEnd and gaolu=:gaolu " + 
                "and mc in (select ylmc from ylfl where ylzl in ('含铁原料')) group by sj) b where a.sj =b.sj order by sj";

            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    
                    hslTable.SetBili(dr.GetDateTime(0), dr.GetString(1), dr.GetDouble(2));
                   
                }
            }
            dr.Close();
            conn.Close();

            foreach (var hslRi in hslTable)
            {
                foreach (var item in hslRi)
                {
                    double hsl = 0;
                    loadRlylHtl(gaolu, item.MingCheng, hslRi.Riqi, out hsl);
                    item.Pinwei = hsl;
                }
            }
    

            DateTime[] riqiAry=new DateTime[hslTable.Count];
            double[] hslAry=new double[hslTable.Count];
            int i=0;
            foreach (var item in hslTable)
	        {
		         riqiAry[i]=item.Riqi;
                hslAry[i]=item.HuiShouLv;
                i++;
            }

            sT = riqiAry;
            s = hslAry;
        }
        // 入炉原料含铁量
        public static void loadRlylHtl(int gaolu, string ylmc, DateTime t, out double s)
        {
            DateTime hist = new DateTime(2014,04,21);
            string yl = "";  //原料
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            double vs = 0;
            string mingcheng = "%4#%";
            Boolean iscyyll = false;  //判断是否是长化验原料矿
            switch (ylmc)
            {
                case "烧结矿":
                    yl = "%烧%";
                    iscyyll = true;
                    break;
                case "竖炉球":
                    yl = "%球%";
                    iscyyll = true;
                    break;
                case "FMG块":
                    if (t < hist)
                    {
                        yl = "%块%";
                        iscyyll = false;

                    }
                    else
                    {
                        yl = "%FMG块%";
                        iscyyll = false;
                    }
                     break;
                  case "PB块":
                     if (t < hist)
                     {
                         yl = "%块%";
                         iscyyll = false;
                     }
                     else
                     {
                         yl = "%PB块%";
                         iscyyll = false;
                     }
                     break;
                default:
                     yl = "%" + ylmc + "%";
                    iscyyll = false;
                    break;
            }

            if (iscyyll == true)
            {

                string sql = "select * from ( select trunc(avg(tfe),2) from ddyl t  where trunc(sj)<=:sj and mc like :mc and cang like :cang   group by trunc(sj) order by trunc(sj) desc ) where rownum =1";
             
               OracleCommand cmd = new OracleCommand(sql,conn);
                cmd.Parameters.Add(":sj", OracleType.DateTime).Value = t;
                cmd.Parameters.Add(":cang", OracleType.NVarChar).Value = mingcheng;
                cmd.Parameters.Add(":mc", OracleType.NVarChar).Value = yl;
                OracleDataReader dr = cmd.ExecuteReader();
 
                while (dr.Read())
                {
                    if (!dr.IsDBNull(0))
                    {

                        vs=dr.GetDouble(0);
                    }
                }
                dr.Close();
            }
            else
            {
                 string sql="";
                 if (ylmc == "PB块" || ylmc == "FMG块")
                 {
                     sql = "select * from (select trunc(tfe,2) from ddyl t  where trunc(sj)<=to_date('" + t.ToString("yyyy-MM-dd") + "','YYYY-MM-DD') and mc like :mc order by sj desc) where rownum=1 ";
                 }
                 else
                 {
                     sql = "select * from (select trunc(tfe,2) from ddyl t  where mc like :mc order by sj desc) where  rownum=1";
                 }
                OracleCommand cmd = new OracleCommand(sql, conn);
                cmd.Parameters.Add(":mc", OracleType.NVarChar).Value = yl;
                OracleDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    if (!dr.IsDBNull(0))
                    {

                        vs = dr.GetDouble(0);
                    }
                }
                dr.Close();
            }
 
            conn.Close();
          
            s = vs;
           
             
                
        }
       
        // 熔剂
        public static void loadRONGJIPeiBi(int gaolu, DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select a.sj,trunc(a.lj*100/b.zong,2) from " +
                        "(select  (nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0)) as lj,sj from rbxiaohao where TRUNC(SJ)>=:sjBegin and TRUNC(SJ)<=:sjEnd and gaolu=:gaolu and mc in (select ylmc from ylfl where ylzl in '熔剂')) a," +
                        "(select sum(nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0)) as zong,sj from rbxiaohao where TRUNC(SJ)>=:sjBegin and TRUNC(SJ)<=:sjEnd and gaolu=:gaolu " +
                        "and mc in (select ylmc from ylfl where ylzl in ('含铁原料','熔剂')) group by sj) b where a.sj =b.sj order by sj";

            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));

                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }
        // 热风温度
        public static void loadRFWD(int gaolu, DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select trunc(sj),trunc(avg(rfwd),2) from rbcaozuo"+
                         " where TRUNC(SJ)>=:sjBegin  and TRUNC(SJ)<=:sjEnd" +
                         " and gaolu=:gaolu group by TRUNC(SJ) order by TRUNC(SJ)";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));

                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }
        // 富氧率
        public static void loadFYL(int gaolu, DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select trunc(sj),trunc(avg((0.78*fyll)/(lfll*60)*100),2) from rbcaozuo" +
                         " where TRUNC(SJ)>=:sjBegin  and TRUNC(SJ)<=:sjEnd" +
                         " and gaolu=:gaolu and lfll<>0 group by TRUNC(SJ) order by TRUNC(SJ)";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));

                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }

        // 实际回收率
        public static void loadSJHSL(int gaolu, DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select a.sj,trunc(a.feliang/b.zong*100,2) from " +
                         "(select trunc(zdsj) as sj,sum(feliang) as feliang from DDluci t where gaolu =:gaolu and trunc(zdsj)>=:sjBegin  and trunc(zdsj)<=:sjEnd group by trunc(zdsj)) a," +
                         "(select sum(nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0)) as zong,sj from rbxiaohao where TRUNC(SJ)>=:sjBegin and TRUNC(SJ)<=:sjEnd and gaolu=:gaolu " +
                          "and  mc in (select ylmc from ylfl where ylzl='含铁原料') group by sj "+
                          ") b where a.sj=b.sj order by a.sj";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));

                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }

        // 5#高炉烧结矿含p负荷
        public static void loadSJKPFH5(int gaolu, DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select c.sj,trunc(c.sjkliang*d.s/100,2) from"+
                        "(select a.sj as sj,trunc(b.zong*1000/a.feliang,2) as sjkliang from "+
                        "(select trunc(zdsj) as sj,sum(feliang) as feliang from DDluci t "+
                         "where gaolu =:gaolu and trunc(zdsj)>=:sjBegin "+
                          "and trunc(zdsj)<=:sjEnd group by trunc(zdsj)) a," +
                          "(select sum(nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0)) as zong,sj from rbxiaohao where TRUNC(SJ)>=:sjBegin "+
                           "and TRUNC(SJ)<=:sjEnd and gaolu =:gaolu " +
                          "and mc like '烧结矿%'  group by sj ) b where a.sj=b.sj order by a.sj) c, "+
                          "(select trunc(h.sj) as sj,trunc(avg(p),6) as s from ddyl h where "+
                          "trunc(h.sj)>=:sjBegin and trunc(h.sj)<:sjEnd and cang like '5#%' " +
                         "and mc like '%大烧%' group by trunc(h.sj)) d "+
                         "where c.sj=d.sj order by c.sj";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));

                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }
        // 1-4#高炉烧结矿含P负荷
        public static void loadSJKPFH(int gaolu, DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select c.sj,trunc(c.sjkliang*d.s/100,2) from" +
                        "(select a.sj as sj,trunc(b.zong*1000/a.feliang,2) as sjkliang from " +
                        "(select trunc(zdsj) as sj,sum(feliang) as feliang from DDluci t " +
                         "where gaolu =:gaolu and trunc(zdsj)>=:sjBegin " +
                          "and trunc(zdsj)<=:sjEnd group by trunc(zdsj)) a," +
                          "(select sum(nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0)) as zong,sj from rbxiaohao where TRUNC(SJ)>=:sjBegin " +
                           "and TRUNC(SJ)<=:sjEnd and gaolu =:gaolu " +
                          "and mc like '烧结矿%'  group by sj ) b where a.sj=b.sj order by a.sj) c, " +
                          "(select trunc(h.sj) as sj,trunc(avg(p),6) as s from ddyl h where " +
                          "trunc(h.sj)>=:sjBegin and trunc(h.sj)<:sjEnd and cang like '1%' " +
                         "and mc like '%大烧%' group by trunc(h.sj)) d " +
                         "where c.sj=d.sj order by c.sj";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));

                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }
        // 5#高炉球团矿p负荷
        public static void loadQTKPFH5(int gaolu, DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            double JTP = 0;

            LoadJTP(out JTP);     //d
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select c.sj,trunc(c.sjkliang*d.s/100,2) from" +
                        "(select a.sj as sj,trunc(b.zong*1000/a.feliang,2) as sjkliang from " +
                        "(select trunc(zdsj) as sj,sum(feliang) as feliang from DDluci t " +
                         "where gaolu =:gaolu and trunc(zdsj)>=:sjBegin " +
                          "and trunc(zdsj)<=:sjEnd group by trunc(zdsj)) a," +
                          "(select sum(nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0)) as zong,sj from rbxiaohao where TRUNC(SJ)>=:sjBegin " +
                           "and TRUNC(SJ)<=:sjEnd and gaolu =:gaolu " +
                          "and mc like '竖%'  group by sj ) b where a.sj=b.sj order by a.sj) c, " +
                          "(select trunc(h.sj) as sj,trunc(avg(p),6) as s from ddyl h where " +
                          "trunc(h.sj)>=:sjBegin and trunc(h.sj)<:sjEnd and cang like '5#%' " +
                         "and mc like '%竖%' group by trunc(h.sj)) d " +
                         "where c.sj=d.sj order by c.sj";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));

                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }
        //读取澳矿的P
        
        public static void loadAOKPFH(int gaolu, DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select c.sj,trunc(c.sjkliang*d.s/100,2) from" +
                        "(select a.sj as sj,trunc(b.zong*1000/a.feliang,2) as sjkliang from " +
                        "(select trunc(zdsj) as sj,sum(feliang) as feliang from DDluci t " +
                         "where gaolu =:gaolu and trunc(zdsj)>=:sjBegin " +
                          "and trunc(zdsj)<=:sjEnd group by trunc(zdsj)) a," +
                          "(select sum(nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0)) as zong,sj from rbxiaohao where TRUNC(SJ)>=:sjBegin " +
                           "and TRUNC(SJ)<=:sjEnd and gaolu =:gaolu " +
                          "and mc like '澳矿%'  group by sj ) b where a.sj=b.sj order by a.sj) c, " +
                          "(select trunc(avg(p),6) as s from ddyl h where " +
                          "mc like '%澳矿%' ) d ";            
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();
          
            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
               
                      dts.Add(dr.GetDateTime(0));
                        vs.Add(dr.GetDouble(1));
                 //  }
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }
        //读取焦炭的P
        public static void LoadJTP(out double JTP)
        {

            JTP = 0;
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection =conn;
             cmd.CommandText = "select * from ddjtsetp where mc='焦炭'";
           OracleDataReader odr = cmd.ExecuteReader();
            while (odr.Read())
            {
               
                if (!(odr[2] is DBNull))
                {
                    JTP = odr.GetDouble(2);    //TFE

                }
            }
            odr.Close();
            conn.Close();
        }
        // 1-4#高炉球团矿p负荷
        public static void loadQTKPFH(int gaolu, DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
           double JTP=0;

           LoadJTP(out JTP);     //d
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            LoadJTP(out JTP);     //d
            string sql = "select c.sj,trunc(c.sjkliang*d.s/100,2) from" +
                        "(select a.sj as sj,trunc(b.zong*1000/a.feliang,2) as sjkliang from " +
                        "(select trunc(zdsj) as sj,sum(feliang) as feliang from DDluci t " +
                         "where gaolu =:gaolu and trunc(zdsj)>=:sjBegin " +
                          "and trunc(zdsj)<=:sjEnd group by trunc(zdsj)) a," +
                          "(select sum(nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0)) as zong,sj from rbxiaohao where TRUNC(SJ)>=:sjBegin " +
                           "and TRUNC(SJ)<=:sjEnd and gaolu =:gaolu " +
                          "and mc like '竖%'  group by sj ) b where a.sj=b.sj order by a.sj) c, " +
                          "(select trunc(h.sj) as sj,trunc(avg(p),6) as s from ddyl h where " +
                          "trunc(h.sj)>=:sjBegin and trunc(h.sj)<:sjEnd and cang like '1%' " +
                         "and mc like '%竖%' group by trunc(h.sj)) d " +
                         "where c.sj=d.sj order by c.sj";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));

                    vs.Add(dr.GetDouble(1) +JTP);
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }




        // 焦比
        public static void loadJB(int gaolu, DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select a.sj,trunc(b.jiaotan/a.feliang*1000,2) from " +
                         "(select trunc(zdsj) as sj,sum(feliang) as feliang from DDluci t where gaolu =:gaolu and trunc(zdsj)>=:sjBegin  and trunc(zdsj)<=:sjEnd group by trunc(zdsj)) a," +
                         "(select sum(BAIBAN)+sum(ZHONGBAN)+sum(YEBAN) as jiaotan,sj from rbxiaohao where TRUNC(SJ)>=:sjBegin and TRUNC(SJ)<=:sjEnd and gaolu=:gaolu " +
                          "and  mc like '焦炭%' group by sj " +
                          ") b where a.sj=b.sj order by a.sj";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));

                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }

        // 煤比
        public static void loadMB(int gaolu, DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select a.sj,trunc(b.jiaotan/a.feliang*1000,2) from " +
                         "(select trunc(zdsj) as sj,sum(feliang) as feliang from DDluci t where gaolu =:gaolu and trunc(zdsj)>=:sjBegin  and trunc(zdsj)<=:sjEnd group by trunc(zdsj)) a," +
                         "(select sum(BAIBAN)+sum(ZHONGBAN)+sum(YEBAN) as jiaotan,sj from rbxiaohao where TRUNC(SJ)>=:sjBegin and TRUNC(SJ)<=:sjEnd and gaolu=:gaolu " +
                          "and  mc like '%煤%' group by sj " +
                          ") b where a.sj=b.sj order by a.sj";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));

                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }
        // 燃料比
        public static void loadRLB(int gaolu, DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select a.sj,trunc(b.jiaotan/a.feliang*1000,2) from " +
                         "(select trunc(zdsj) as sj,sum(feliang) as feliang from DDluci t where gaolu =:gaolu and trunc(zdsj)>=:sjBegin  and trunc(zdsj)<=:sjEnd group by trunc(zdsj)) a," +
                         "(select sum(BAIBAN)+sum(ZHONGBAN)+sum(YEBAN) as jiaotan,sj from rbxiaohao where TRUNC(SJ)>=:sjBegin and TRUNC(SJ)<=:sjEnd and gaolu=:gaolu " +
                          "and  (mc like '%煤%' or mc like '焦%') group by sj " +
                          ") b where a.sj=b.sj order by a.sj";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));

                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }

        // 热风压力
        public static void loadRFYL(int gaolu, DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select trunc(sj),trunc(avg(rfyl),2) from rbcaozuo "+
                         "where TRUNC(SJ)>=:sjBegin and TRUNC(SJ)<=:sjEnd " +
                         "and gaolu=:gaolu group by trunc(sj) order by trunc(sj)";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));

                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }

        // 炉顶压力
        public static void loadLDYL(int gaolu, DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select trunc(sj),trunc(avg(ldyl),2) from rbcaozuo " +
                         "where TRUNC(SJ)>=:sjBegin and TRUNC(SJ)<=:sjEnd " +
                         "and gaolu=:gaolu group by trunc(sj) order by trunc(sj) ";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));

                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }
        //日产量
        public static void loadRcl(int gaolu, DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select trunc(zdsj),sum(feliang) from ddluci   " +
                         "where trunc(zdsj)>=:sjBegin and trunc(zdsj)<=:sjEnd " +
                         "and gaolu=:gaolu group by trunc(zdsj) order by trunc(zdsj) ";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));

                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }


        //炉渣碱度
        public static void loadlzjd(int gaolu, DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select trunc(zdsj),trunc(avg(zhar2),2) from ddluci   " +
                         "where trunc(zdsj)>=:sjBegin and trunc(zdsj)<=:sjEnd " +
                         "and gaolu=:gaolu group by trunc(zdsj) order by trunc(zdsj) ";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));

                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }
        //铁中SI
        public static void loadFESI(int gaolu, DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select trunc(zdsj),trunc(avg(FESI),2) from ddluci   " +
                         "where trunc(zdsj)>=:sjBegin and trunc(zdsj)<=:sjEnd " +
                         "and gaolu=:gaolu group by trunc(zdsj) order by trunc(zdsj) ";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                if (!dr.IsDBNull(1))
                {
                    dts.Add(dr.GetDateTime(0));

                    vs.Add(dr.GetDouble(1));
                }
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }
        //高炉消耗含铁原料
        public static void loadHTYL (int gaolu, DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
    {
        OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
        conn.Open();
        List<DateTime> dts = new List<DateTime>();
        List<double> vs = new List<double>();
        string sql = " select sum(nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0)) as zong,sj from rbxiaohao" +
        " where TRUNC(SJ)>=:sjBegin and TRUNC(SJ)<=:sjEnd and gaolu=:gaolu" +
       " and  mc in (select ylmc from ylfl where ylzl='含铁原料') group by sj order by sj";
        OracleCommand cmd = new OracleCommand(sql, conn);
        cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
        cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
        cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
        OracleDataReader dr = cmd.ExecuteReader();
        while (dr.Read())
        {
            if (!dr.IsDBNull(1))
            {
                vs.Add(dr.GetDouble(0));
                dts.Add(dr.GetDateTime(1));
            }
        }
        dr.Close();
        conn.Close();
        sT = dts.ToArray();
        s = vs.ToArray();
    
    }
    
    }
}
