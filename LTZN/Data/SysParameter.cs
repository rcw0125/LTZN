using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;

namespace LTZN.Data
{
    /// <summary>
    /// 系统参数
    /// </summary>
    public class SysParameter
    {
        public static void SetParameter(string paraName,double val)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string checkSql = "select count(*) from LTZN_Para where Name=:name";
            OracleCommand checkCmd = new OracleCommand(checkSql, conn);
            checkCmd.Parameters.Add(":name", OracleType.VarChar).Value = paraName;
            int count = Convert.ToInt32(checkCmd.ExecuteScalar());

            if (count > 0)
            {
                string sql = "update LTZN_Para set Val=:vl where Name=:name";
                OracleCommand saveCmd = new OracleCommand(sql, conn);
                OracleParameter valPara = saveCmd.Parameters.Add(":vl", OracleType.Double);
                OracleParameter namePara = saveCmd.Parameters.Add(":name", OracleType.VarChar);

                namePara.Value =paraName;
                valPara.Value = val;
                saveCmd.ExecuteNonQuery();
            }
            else
            {
                string sql = "insert into LTZN_Para(Name,Val) Values(:name,:vl)";
                OracleCommand saveCmd = new OracleCommand(sql, conn);
                OracleParameter namePara = saveCmd.Parameters.Add(":name", OracleType.VarChar);
                OracleParameter valPara = saveCmd.Parameters.Add(":vl", OracleType.Double);
                namePara.Value = paraName;
                valPara.Value = val;
                saveCmd.ExecuteNonQuery();
            }
            conn.Close();
        }

        public static void SetParameter(string paraName, string val)
        {
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string checkSql = "select count(*) from LTZN_Para where Name=:name";
            OracleCommand checkCmd = new OracleCommand(checkSql, conn);
            checkCmd.Parameters.Add(":name", OracleType.VarChar).Value=paraName;
            int count = Convert.ToInt32(checkCmd.ExecuteScalar());

            if (count > 0)
            {
                string sql = "update LTZN_Para set STRVAL=:vl where Name=:name";
                OracleCommand saveCmd = new OracleCommand(sql, conn);
                saveCmd.Parameters.Add(":vl", OracleType.VarChar, 1000).Value = val;     
                saveCmd.Parameters.Add(":name", OracleType.VarChar, 1000).Value = paraName;
                saveCmd.ExecuteNonQuery();
            }
            else
            {
                string sql = "insert into LTZN_Para(Name,STRVAL) Values(:name,:vl)";
                OracleCommand saveCmd = new OracleCommand(sql, conn);
                saveCmd.Parameters.Add(":name", OracleType.VarChar, 1000).Value = paraName;
                saveCmd.Parameters.Add(":vl", OracleType.VarChar, 1000).Value = val;
                saveCmd.ExecuteNonQuery();
            }
            conn.Close();
        }

        public static string GetStrParameter(string paraName)
        {
            string result="";
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string checkSql = "select STRVAL from LTZN_Para where Name=:name";
            OracleCommand checkCmd = new OracleCommand(checkSql, conn);
            checkCmd.Parameters.Add(":name", OracleType.VarChar).Value = paraName;
            OracleDataReader dr = checkCmd.ExecuteReader();
            if (dr.Read())
            {
                if (!dr.IsDBNull(0)) result = dr.GetString(0);

            }
            dr.Close();
            conn.Close();
            return result;

        }

        public static string GetStrParameter(string paraName,string defaultVal)
        {
            string result = "";
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string checkSql = "select STRVAL from LTZN_Para where Name=:name";
            OracleCommand checkCmd = new OracleCommand(checkSql, conn);
            checkCmd.Parameters.Add(":name", OracleType.VarChar).Value = paraName;
            OracleDataReader dr = checkCmd.ExecuteReader();
            if (dr.Read())
            {
                if (!dr.IsDBNull(0)) result = dr.GetString(0);
            }
            dr.Close();
            conn.Close();
            if (result == "")
            {
                SetParameter(paraName, defaultVal);
                result = defaultVal;
            }
            return result;
        }


        public static double GetDblParameter(string paraName)
        {
            double result =0;
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string checkSql = "select VAL from LTZN_Para where Name=:name";
            OracleCommand checkCmd = new OracleCommand(checkSql, conn);
            checkCmd.Parameters.Add(":name", OracleType.VarChar).Value = paraName;
            OracleDataReader dr = checkCmd.ExecuteReader();
            if (dr.Read())
            {
                if (!dr.IsDBNull(0)) result = dr.GetDouble(0);
              
            }
            dr.Close();
            conn.Close();
            return result;
        }

        public static double GetDblParameter(string paraName,double defaultVal)
        {
            double result = 0;
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string checkSql = "select VAL from LTZN_Para where Name=:name";
            OracleCommand checkCmd = new OracleCommand(checkSql, conn);
            checkCmd.Parameters.Add(":name", OracleType.VarChar).Value = paraName;
            OracleDataReader dr = checkCmd.ExecuteReader();
            if (dr.Read())
            {
                if (!dr.IsDBNull(0)) result = dr.GetDouble(0);

            }
            dr.Close();
            conn.Close();
            if (result == 0)
            {
                SetParameter(paraName, defaultVal);
                result = defaultVal;
            }
            return result;
        }
    
    }
}
