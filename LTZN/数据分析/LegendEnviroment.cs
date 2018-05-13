using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;

namespace LTZN.数据分析
{
    class LegendEnviroment
    {
       public  static void loadSYR(int gaolu, DateTime t1, DateTime t2, out DateTime[] shengyureT, out double[] shengyuere, out DateTime[] yurebiliT, out double[] yurebili, out DateTime[] siT, out double[] si)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();
            List<DateTime> dts1 = new List<DateTime>();
            List<double> vs1 = new List<double>();
            List<double> vs2 = new List<double>();

            string sql = "select dksj,FESi from ddluci where dksj>=:sjBegin and dksj<=:sjEnd and gaolu=:gaolu and FeSi is not null order by dksj";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                dts.Add(dr.GetDateTime(0));
                vs.Add(dr.GetDouble(1));
            }
            dr.Close();

            sql = "select T,SYR,SYRBL from LT_LIAO where T>=:sjBegin and T<=:sjEnd and gaolu=:gaolu and syr is not null and syrbl is not null order by t";
            cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            dr = cmd.ExecuteReader();

            while (dr.Read())
            {  
                dts1.Add(dr.GetDateTime(0));
                vs1.Add(dr.GetDouble(1));
                vs2.Add(dr.GetDouble(2));
            }
            dr.Close();

            conn.Close();
            shengyureT = dts1.ToArray();
            yurebiliT = dts1.ToArray();
            siT = dts.ToArray();
            shengyuere = vs1.ToArray();
            yurebili = vs2.ToArray();
            si = vs.ToArray();
        }

        public static void loadSi(int gaolu, DateTime t1, DateTime t2, out DateTime[] siT, out double[] si)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select dksj,FESi from ddluci where dksj>=:sjBegin and dksj<=:sjEnd and gaolu=:gaolu and FeSi is not null order by dksj";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                dts.Add(dr.GetDateTime(0));
                vs.Add(dr.GetDouble(1));
            }
            dr.Close();
            conn.Close();
            siT = dts.ToArray();
            si = vs.ToArray();
        }
        public static void loadSi(int gaolu, string luci1, string luci2, out long[] siLuci, out double[] si)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<long> dts = new List<long>();
            List<double> vs = new List<double>();

            string sql = "select luci,FESi from ddluci where luci>=:luciBegin and luci<=:luciEnd and gaolu=:gaolu and FeSi is not null and luci is not null order by luci";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":luciBegin", OracleType.VarChar).Value = luci1;
            cmd.Parameters.Add(":luciEnd", OracleType.VarChar).Value = luci2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                dts.Add(Convert.ToInt64(dr.GetString(0)));
                vs.Add(dr.GetDouble(1));
            }
            dr.Close();
            conn.Close();
            siLuci = dts.ToArray();
            si = vs.ToArray();
        }
        public static void loadSiSetting(int gaolu, out double[] dy, out double max, out double min)
        {

            dy = new double[4];
            max = 1;
            min = 0;
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string sql = "select name,val from LTZN_Para";
            OracleCommand cmd = new OracleCommand(sql, conn);
            OracleDataReader dr = cmd.ExecuteReader();
            string[] paraName = new string[] { gaolu.ToString() + "高炉Si最大", gaolu.ToString() + "高炉Si次大", gaolu.ToString() + "高炉Si次小", gaolu.ToString() + "高炉Si最小", "Si显示最大值", "Si显示最小值" };


            while (dr.Read())
            {
                string name = dr.GetString(0);
                if (!dr.IsDBNull(1))
                {
                    for (int i = 0; i < 6; i++)
                    {
                        if (name == paraName[i])
                        {
                            if (i == 4)
                                max = dr.GetDouble(1);
                            else
                            {
                                if (i == 5)
                                    min = dr.GetDouble(1);
                                else
                                    dy[i] = dr.GetDouble(1);
                            }
                        }
                    }
                }
            }
                dr.Close();
                conn.Close(); 
        }
        public static double stddev(double[] si)
        {
            if (si.Length<=0)
                return 0;
            double sumx = 0;
            for (int i = 0; i < si.Length; i++)
                sumx += si[i];
            double avgx = sumx / si.Length;
            double sumx2 = 0;
            for (int i = 0; i < si.Length; i++)
                sumx2 += Math.Pow(si[i] - avgx, 2);
            return Math.Sqrt(sumx2 / si.Length);
        }
        public static void loadWlrSetting(int gaolu, out double[] dy, out double max, out double min)
        {

            dy = new double[4];
            max = 1;
            min = 0;
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string sql = "select name,val from LTZN_Para";
            OracleCommand cmd = new OracleCommand(sql, conn);
            OracleDataReader dr = cmd.ExecuteReader();
            string[] paraName = new string[] { gaolu.ToString() + "高炉物理热最大", gaolu.ToString() + "高炉物理热次大", gaolu.ToString() + "高炉物理热次小", gaolu.ToString() + "高炉物理热最小", "物理热显示最大值", "物理热显示最小值" };


            while (dr.Read())
            {
                string name = dr.GetString(0);
                if (!dr.IsDBNull(1))
                {
                    for (int i = 0; i < 6; i++)
                    {
                        if (name == paraName[i])
                        {
                            if (i == 4)
                                max = dr.GetDouble(1);
                            else
                            {
                                if (i == 5)
                                    min = dr.GetDouble(1);
                                else
                                    dy[i] = dr.GetDouble(1);
                            }
                        }
                    }
                }
            }
            dr.Close();
            conn.Close();
        }
        public static void newloadSSetting(int gaolu, out double[] dy, out double max, out double min)
        {

            dy = new double[4];
            max = 1;
            min = 0;
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string sql = "select name,val from LTZN_Para";
            OracleCommand cmd = new OracleCommand(sql, conn);
            OracleDataReader dr = cmd.ExecuteReader();
            string[] paraName = new string[] { gaolu.ToString() + "高炉S最大", gaolu.ToString() + "高炉S次大", gaolu.ToString() + "高炉S次小", gaolu.ToString() + "高炉S最小", "S显示最大值", "S显示最小值" };


            while (dr.Read())
            {
                string name = dr.GetString(0);
                if (!dr.IsDBNull(1))
                {
                    for (int i = 0; i < 6; i++)
                    {
                        if (name == paraName[i])
                        {
                            if (i == 4)
                                max = dr.GetDouble(1);
                            else
                            {
                                if (i == 5)
                                    min = dr.GetDouble(1);
                                else
                                    dy[i] = dr.GetDouble(1);
                            }
                        }
                    }
                }
            }
            dr.Close();
            conn.Close();
        }
        public static void newloadRSetting(int gaolu, out double[] dy, out double max, out double min)
        {

            dy = new double[4];
            max = 1;
            min = 0;
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string sql = "select name,val from LTZN_Para";
            OracleCommand cmd = new OracleCommand(sql, conn);
            OracleDataReader dr = cmd.ExecuteReader();
            string[] paraName = new string[] { gaolu.ToString() + "高炉R最大", gaolu.ToString() + "高炉R次大", gaolu.ToString() + "高炉R次小", gaolu.ToString() + "高炉R最小", "R显示最大", "R显示最小" };


            while (dr.Read())
            {
                string name = dr.GetString(0);
                if (!dr.IsDBNull(1))
                {
                    for (int i = 0; i < 6; i++)
                    {
                        if (name == paraName[i])
                        {
                            if (i == 4)
                                max = dr.GetDouble(1);
                            else
                            {
                                if (i == 5)
                                    min = dr.GetDouble(1);
                                else
                                    dy[i] = dr.GetDouble(1);
                            }
                        }
                    }
                }
            }
            dr.Close();
            conn.Close();
        }
        public static void loadR(int gaolu, DateTime t1, DateTime t2, out DateTime[] RT, out double[] R)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select dksj,ZHAR2 from ddluci where dksj>=:sjBegin and dksj<=:sjEnd and gaolu=:gaolu and ZHAR2 is not null order by dksj";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                dts.Add(dr.GetDateTime(0));
                vs.Add(dr.GetDouble(1));
            }
            dr.Close();
            conn.Close();
            RT = dts.ToArray();
            R = vs.ToArray();
        }
        public static void loadR(int gaolu, string luci1, string luci2, out long[] RLuci, out double[] R)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<long> dts = new List<long>();
            List<double> vs = new List<double>();

            string sql = "select luci,ZHAR2 from ddluci where luci>=:luciBegin and luci<=:luciEnd and gaolu=:gaolu and ZHAR2 is not null and luci is not null order by luci";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":luciBegin", OracleType.VarChar).Value = luci1;
            cmd.Parameters.Add(":luciEnd", OracleType.VarChar).Value = luci2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                dts.Add(Convert.ToInt64(dr.GetString(0)));
                vs.Add(dr.GetDouble(1));
            }
            dr.Close();
            conn.Close();
            RLuci = dts.ToArray();
            R = vs.ToArray();
        }
        public static void loadCanShuSetting(int gaolu,string canshu, out double[] dy, out double max, out double min)
        {

            dy = new double[2];
            max = 1;
            min = 0;
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string sql = "select * from LT_QSPara where canshu=:canshu and gaolu=:gaolu";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":canshu", OracleType.VarChar).Value = canshu;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                string name = dr.GetString(0);
                if (!dr.IsDBNull(1))
                {
                 
                            dy[0] = dr.GetDouble(1);
                            dy[1] = dr.GetDouble(2);
                            max = dr.GetDouble(3);
                            min = dr.GetDouble(4);
                }
            }
            dr.Close();
            conn.Close();
        }
        public static void loadRSetting(out double[] dy, out double max, out double min)
        {

            dy = new double[4];
            max = 1;
            min = 0;
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string sql = "select name,val from LTZN_Para";
            OracleCommand cmd = new OracleCommand(sql, conn);
            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                string name = dr.GetString(0);
                if (!dr.IsDBNull(1))
                {
                    switch (name)
                    {
                        case "R最大":
                            dy[0] = dr.GetDouble(1);
                            break;
                        case "R次大":
                            dy[1] = dr.GetDouble(1);
                            break;
                        case "R次小":
                            dy[2] = dr.GetDouble(1);
                            break;
                        case "R最小":
                            dy[3] = dr.GetDouble(1);
                            break;
                        case "R显示最大值":
                            max = dr.GetDouble(1);
                            break;
                        case "R显示最小值":
                            min = dr.GetDouble(1);
                            break;
                    }
                }
            }
            dr.Close();
            conn.Close();
        }
        public static void loadTiO2(int gaolu,string canshu, DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select dksj,"+canshu +" from ddluci where dksj>=:sjBegin and dksj<=:sjEnd and gaolu=:gaolu and "+canshu+" is not null order by dksj";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                dts.Add(dr.GetDateTime(0));
                vs.Add(dr.GetDouble(1));
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }

        public static void loadTiO2(int gaolu,string canshu, string luci1, string luci2, out long[] sLuci, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<long> dts = new List<long>();
            List<double> vs = new List<double>();

            string sql = "select luci,"+canshu+" from ddluci where luci>=:luciBegin and luci<=:luciEnd and gaolu=:gaolu and "+canshu+" is not null and luci is not null order by luci";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":luciBegin", OracleType.VarChar).Value = luci1;
            cmd.Parameters.Add(":luciEnd", OracleType.VarChar).Value = luci2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                dts.Add(Convert.ToInt64(dr.GetString(0)));
                vs.Add(dr.GetDouble(1));
            }
            dr.Close();
            conn.Close();
            sLuci = dts.ToArray();
            s = vs.ToArray();
        }

        public static void loadshaojie(int gaolu, string canshu, DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {   
            string shaojieji="";
            if (gaolu==7) 
                shaojieji="1#大烧";
            else if (gaolu==8) shaojieji="2#大烧";
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();

            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select sj," + canshu + " from ddyl where sj>=:sjBegin and sj<=:sjEnd and MC=:gaolu and " + canshu + " is not null order by sj";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.NVarChar).Value = shaojieji;
            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                dts.Add(dr.GetDateTime(0));
                vs.Add(dr.GetDouble(1));
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();

        }





        public static void loadS(int gaolu, DateTime t1, DateTime t2, out DateTime[] sT, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select dksj,FES from ddluci where dksj>=:sjBegin and dksj<=:sjEnd and gaolu=:gaolu and FeS is not null order by dksj";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                dts.Add(dr.GetDateTime(0));
                vs.Add(dr.GetDouble(1));
            }
            dr.Close();
            conn.Close();
            sT = dts.ToArray();
            s = vs.ToArray();
        }
        public static void loadS(int gaolu, string luci1, string luci2, out long[] sLuci, out double[] s)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<long> dts = new List<long>();
            List<double> vs = new List<double>();

            string sql = "select luci,FES from ddluci where luci>=:luciBegin and luci<=:luciEnd and gaolu=:gaolu and FeS is not null and luci is not null order by luci";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":luciBegin", OracleType.VarChar).Value = luci1;
            cmd.Parameters.Add(":luciEnd", OracleType.VarChar).Value = luci2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                dts.Add(Convert.ToInt64(dr.GetString(0)));
                vs.Add(dr.GetDouble(1));
            }
            dr.Close();
            conn.Close();
            sLuci = dts.ToArray();
            s = vs.ToArray();
        }
        public static void loadSSetting(out double[] dy1, out double[] dy2, out double max, out double min,out double max2,out double min2)
        {

            dy1 = new double[3];
            dy2 = new double[3];
            max = 0.05;
            min = 0;
            max2 = 0.02;
            min2 = 0;
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string sql = "select name,val from LTZN_Para";
            OracleCommand cmd = new OracleCommand(sql, conn);
            OracleDataReader dr = cmd.ExecuteReader();
            string[] paraName = new string[] { "S单值上限", "S单值均值", "S单值下限", "S极差上限", "S极差均值", "S极差下限", "S单值图显示最大值", "S单值图显示最小值", "S极差图显示最大值", "S极差图显示最小值" };


            while (dr.Read())
            {
                string name = dr.GetString(0);
                if (!dr.IsDBNull(1))
                {
                    for (int i = 0; i < 10; i++)
                    {
                        if (name == paraName[i])
                        {
                            if (i == 6)
                                max = dr.GetDouble(1);
                            else
                            {
                                if (i == 7)
                                    min = dr.GetDouble(1);
                                else
                                {
                                    if (i == 8)
                                        max2 = dr.GetDouble(1);
                                    else
                                    {
                                        if (i == 9)
                                            min2 = dr.GetDouble(1);
                                        else
                                        {
                                            if (i < 3)
                                                dy1[i] = dr.GetDouble(1);
                                            else
                                                dy2[i - 3] = dr.GetDouble(1);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            dr.Close();
            conn.Close();
        }


        public static void loadLiang(int gaolu, DateTime t1, DateTime t2, out DateTime[] liangT, out double[] liang)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();

            string sql = "select dksj,FELiang from ddluci where dksj>=:sjBegin and dksj<=:sjEnd and gaolu=:gaolu and FeLiang is not null order by dksj";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = t1;
            cmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = t2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                dts.Add(dr.GetDateTime(0));
                vs.Add(dr.GetDouble(1));
            }
            dr.Close();
            conn.Close();
            liangT = dts.ToArray();
            liang = vs.ToArray();
        }
        public static void loadLiang(int gaolu, string luci1, string luci2, out long[] liangLuci, out double[] liang)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            List<long> dts = new List<long>();
            List<double> vs = new List<double>();

            string sql = "select luci,FELiang from ddluci where luci>=:luciBegin and luci<=:luciEnd and gaolu=:gaolu and FeLiang is not null and luci is not null order by luci";
            OracleCommand cmd = new OracleCommand(sql, conn);
            cmd.Parameters.Add(":luciBegin", OracleType.VarChar).Value = luci1;
            cmd.Parameters.Add(":luciEnd", OracleType.VarChar).Value = luci2;
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmd.ExecuteReader();

            while (dr.Read())
            {
                dts.Add(Convert.ToInt64(dr.GetString(0)));
                vs.Add(dr.GetDouble(1));
            }
            dr.Close();
            conn.Close();
            liangLuci = dts.ToArray();
            liang = vs.ToArray();
        }
        public static void loadLiangSetting(int gaolu, out double[] dy, out double max, out double min)
        {

            dy = new double[4];
            max = 200;
            min = 0;
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string sql = "select name,val from LTZN_Para";
            OracleCommand cmd = new OracleCommand(sql, conn);
            OracleDataReader dr = cmd.ExecuteReader();
            string[] paraName = new string[] { gaolu.ToString() + "高炉Liang最大", gaolu.ToString() + "高炉Liang次大", gaolu.ToString() + "高炉Liang次小", gaolu.ToString() + "高炉Liang最小", "Liang显示最大值", "Liang显示最小值" };


            while (dr.Read())
            {
                string name = dr.GetString(0);
                if (!dr.IsDBNull(1))
                {
                    for (int i = 0; i < 6; i++)
                    {
                        if (name == paraName[i])
                        {
                            if (i == 4)
                                max = dr.GetDouble(1);
                            else
                            {
                                if (i == 5)
                                    min = dr.GetDouble(1);
                                else
                                    dy[i] = dr.GetDouble(1);
                            }
                        }
                    }
                }
            }
            dr.Close();
            conn.Close();
        }
    
    }
}
