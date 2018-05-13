using System;
using System.Collections.Generic;
using System.Text;
using LTZN.Data;
using System.Data.OracleClient;
using System.Data;

namespace LTZN.Repository
{
    public class MonthEnergyRepository : RepositoryBase
    {
        public void Save(MonthEnergy entity)
        {
            conn.Open();
            OracleCommand cmd1 = new OracleCommand();
            cmd1.Connection = conn;
            cmd1.CommandText = "Select count(*) from 每月能耗 Where 年=:nian and 月=:yue";
            cmd1.Parameters.Add(":nian", OracleType.Int32).Value = entity.Nian;
            cmd1.Parameters.Add(":yue", OracleType.Int32).Value = entity.Yue;
            int count = Convert.ToInt32(cmd1.ExecuteScalar());
            if (count == 0)
            {
                OracleCommand insertCmd = new OracleCommand();
                insertCmd.Connection = conn;
                insertCmd.CommandText = "INSERT INTO 每月能耗(年,月,耗电量,蒸汽,耗水,高炉煤气,煤气回收,富氧量,氮气,大块铁,高炉1人数,高炉2人数,高炉3人数,高炉4人数,高炉5人数,高炉6人数) VALUES(:年,:月,:耗电量,:蒸汽,:耗水,:高炉煤气,:煤气回收,:富氧量,:氮气,:大块铁,:高炉1人数,:高炉2人数,:高炉3人数,:高炉4人数,:高炉5人数,:高炉6人数)";

                insertCmd.Parameters.Add(":年", OracleType.Int32).Value = entity.Nian;
                insertCmd.Parameters.Add(":月", OracleType.Int32).Value = entity.Yue;

                insertCmd.Parameters.Add(":耗电量", OracleType.Double, 22).Value = (object)entity.HaoDianLiang ?? DBNull.Value;
                insertCmd.Parameters.Add(":蒸汽", OracleType.Double, 22).Value = (object)entity.ZhengQi ?? DBNull.Value;
                insertCmd.Parameters.Add(":耗水", OracleType.Double, 22).Value = (object)entity.HaoShui ?? DBNull.Value;
                insertCmd.Parameters.Add(":高炉煤气", OracleType.Double, 22).Value = (object)entity.GaoLuMeiQi ?? DBNull.Value;
                insertCmd.Parameters.Add(":煤气回收", OracleType.Double, 22).Value = (object)entity.HuiShouMeiQi ?? DBNull.Value;
                insertCmd.Parameters.Add(":富氧量", OracleType.Double, 22).Value = (object)entity.FuYang ?? DBNull.Value;
                insertCmd.Parameters.Add(":氮气", OracleType.Double, 22).Value = (object)entity.DanQi ?? DBNull.Value;
                insertCmd.Parameters.Add(":大块铁", OracleType.Double, 22).Value = (object)entity.DaKuaiTie ?? DBNull.Value;
                insertCmd.Parameters.Add(":高炉1人数", OracleType.Int32).Value = (object)entity.GaoLu1RenShu ?? DBNull.Value;
                insertCmd.Parameters.Add(":高炉2人数", OracleType.Int32).Value = (object)entity.GaoLu2RenShu ?? DBNull.Value;
                insertCmd.Parameters.Add(":高炉3人数", OracleType.Int32).Value = (object)entity.GaoLu3RenShu ?? DBNull.Value;
                insertCmd.Parameters.Add(":高炉4人数", OracleType.Int32).Value = (object)entity.GaoLu4RenShu ?? DBNull.Value;
                insertCmd.Parameters.Add(":高炉5人数", OracleType.Int32).Value = (object)entity.GaoLu5RenShu ?? DBNull.Value;
                insertCmd.Parameters.Add(":高炉6人数", OracleType.Int32).Value = (object)entity.GaoLu6RenShu ?? DBNull.Value;

                insertCmd.ExecuteNonQuery();
            }
            else
            {
                OracleCommand updateCmd = new OracleCommand();
                updateCmd.Connection = conn;
                updateCmd.CommandText = "UPDATE 每月能耗 SET 耗电量=:耗电量,蒸汽=:蒸汽,耗水=:耗水,高炉煤气=:高炉煤气,煤气回收=:煤气回收,富氧量=:富氧量,氮气=:氮气,大块铁=:大块铁,高炉1人数=:高炉1人数,高炉2人数=:高炉2人数,高炉3人数=:高炉3人数,高炉4人数=:高炉4人数,高炉5人数=:高炉5人数,高炉6人数=:高炉6人数 WHERE 年=:年 and 月=:月";

                updateCmd.Parameters.Add(":耗电量", OracleType.Double, 22).Value = (object)entity.HaoDianLiang ?? DBNull.Value;
                updateCmd.Parameters.Add(":蒸汽", OracleType.Double, 22).Value = (object)entity.ZhengQi ?? DBNull.Value;
                updateCmd.Parameters.Add(":耗水", OracleType.Double, 22).Value = (object)entity.HaoShui ?? DBNull.Value;
                updateCmd.Parameters.Add(":高炉煤气", OracleType.Double, 22).Value = (object)entity.GaoLuMeiQi ?? DBNull.Value;
                updateCmd.Parameters.Add(":煤气回收", OracleType.Double, 22).Value = (object)entity.HuiShouMeiQi ?? DBNull.Value;
                updateCmd.Parameters.Add(":富氧量", OracleType.Double, 22).Value = (object)entity.FuYang ?? DBNull.Value;
                updateCmd.Parameters.Add(":氮气", OracleType.Double, 22).Value = (object)entity.DanQi ?? DBNull.Value;
                updateCmd.Parameters.Add(":大块铁", OracleType.Double, 22).Value = (object)entity.DaKuaiTie ?? DBNull.Value;
                updateCmd.Parameters.Add(":高炉1人数", OracleType.Int32).Value = (object)entity.GaoLu1RenShu ?? DBNull.Value;
                updateCmd.Parameters.Add(":高炉2人数", OracleType.Int32).Value = (object)entity.GaoLu2RenShu ?? DBNull.Value;
                updateCmd.Parameters.Add(":高炉3人数", OracleType.Int32).Value = (object)entity.GaoLu3RenShu ?? DBNull.Value;
                updateCmd.Parameters.Add(":高炉4人数", OracleType.Int32).Value = (object)entity.GaoLu4RenShu ?? DBNull.Value;
                updateCmd.Parameters.Add(":高炉5人数", OracleType.Int32).Value = (object)entity.GaoLu5RenShu ?? DBNull.Value;
                updateCmd.Parameters.Add(":高炉6人数", OracleType.Int32).Value = (object)entity.GaoLu6RenShu ?? DBNull.Value;

                updateCmd.Parameters.Add(":年", OracleType.Int32).Value = entity.Nian;
                updateCmd.Parameters.Add(":月", OracleType.Int32).Value = entity.Yue;
                updateCmd.ExecuteNonQuery();
            }

            OracleCommand cmd2 = new OracleCommand();
            cmd2.Connection = conn;
            cmd2.CommandText = "Select count(*) from 每月煤 Where 年=:nian and 月=:yue";
            cmd2.Parameters.Add(":nian", OracleType.Int32).Value = entity.Nian;
            cmd2.Parameters.Add(":yue", OracleType.Int32).Value = entity.Yue;
            if (Convert.ToInt32(cmd2.ExecuteScalar()) == 0)
            {
                OracleCommand insertCmd = new OracleCommand();
                insertCmd.Connection = conn;
                insertCmd.CommandText = "INSERT INTO 每月煤(年,月,重量,水份,灰份,挥发份,S) VALUES(:年,:月,:重量,:水份,:灰份,:挥发份,:S)";

                insertCmd.Parameters.Add(":年", OracleType.Int32).Value = entity.Nian;
                insertCmd.Parameters.Add(":月", OracleType.Int32).Value = entity.Yue;

                insertCmd.Parameters.Add(":重量", OracleType.Double, 22).Value = (object)entity.ShiMeiFen ?? DBNull.Value;
                insertCmd.Parameters.Add(":水份", OracleType.Double, 22).Value = (object)entity.MeiShuiFen ?? DBNull.Value;
                insertCmd.Parameters.Add(":灰份", OracleType.Double, 22).Value = (object)entity.MeiHuiFen ?? DBNull.Value;
                insertCmd.Parameters.Add(":挥发份", OracleType.Double, 22).Value = (object)entity.MeiHuiFaFen ?? DBNull.Value;
                insertCmd.Parameters.Add(":S", OracleType.Double, 22).Value = (object)entity.MeiS ?? DBNull.Value;

                insertCmd.ExecuteNonQuery();
            }
            else
            {
                OracleCommand updateCmd = new OracleCommand();
                updateCmd.Connection = conn;
                updateCmd.CommandText = "UPDATE 每月煤 SET 重量=:重量,水份=:水份,灰份=:灰份,挥发份=:挥发份,S=:S WHERE 年=:年 and 月=:月";

                updateCmd.Parameters.Add(":重量", OracleType.Double, 22).Value = (object)entity.ShiMeiFen ?? DBNull.Value;
                updateCmd.Parameters.Add(":水份", OracleType.Double, 22).Value = (object)entity.MeiShuiFen ?? DBNull.Value;
                updateCmd.Parameters.Add(":灰份", OracleType.Double, 22).Value = (object)entity.MeiHuiFen ?? DBNull.Value;
                updateCmd.Parameters.Add(":挥发份", OracleType.Double, 22).Value = (object)entity.MeiHuiFaFen ?? DBNull.Value;
                updateCmd.Parameters.Add(":S", OracleType.Double, 22).Value = (object)entity.MeiS ?? DBNull.Value;

                updateCmd.Parameters.Add(":年", OracleType.Int32).Value = entity.Nian;
                updateCmd.Parameters.Add(":月", OracleType.Int32).Value = entity.Yue;

                updateCmd.ExecuteNonQuery();
            }
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "UPDATE每月焦";
            cmd.Parameters.Clear();
            cmd.Parameters.Add("nian", OracleType.Int32).Value = entity.Nian;
            cmd.Parameters.Add("yue", OracleType.Int32).Value = entity.Yue;
            cmd.Parameters.Add("ganmaojiao", OracleType.Number).Value = (object)entity.GanMaoJiao ?? DBNull.Value;
            cmd.Parameters.Add("shijiaofen", OracleType.Number).Value = (object)entity.ShiJiaoFen ?? DBNull.Value;
            cmd.Parameters.Add("ziyongjiaofen", OracleType.Number).Value = (object)entity.JiaoDing ?? DBNull.Value;
            cmd.Parameters.Add("jiaotanshuifen", OracleType.Number).Value = (object)entity.JiaoShuiFen ?? DBNull.Value;
            cmd.Parameters.Add("jiaotanhuifen", OracleType.Number).Value = (object)entity.JiaoHuiFen ?? DBNull.Value;
            cmd.Parameters.Add("jiaotanhuifafen", OracleType.Number).Value = (object)entity.JiaoHuiFaFen ?? DBNull.Value;
            cmd.Parameters.Add("jiaotans", OracleType.Number).Value = (object)entity.JiaoS ?? DBNull.Value;
            cmd.Parameters.Add("jiaotanc", OracleType.Number).Value = (object)entity.JiaoC ?? DBNull.Value;
            cmd.Parameters.Add("jiaotanm25", OracleType.Number).Value = (object)entity.JiaoM25 ?? DBNull.Value;
            cmd.Parameters.Add("jiaotanm10", OracleType.Number).Value = (object)entity.JiaoM10 ?? DBNull.Value;
            cmd.ExecuteNonQuery();

            conn.Close();
        }

        public MonthEnergy GetMonthEnergy(int nian, int yue)
        {
            conn.Open();
            MonthEnergy entity = new MonthEnergy();

            entity.Nian = nian;
            entity.Yue = yue;

            OracleCommand selectCmd = new OracleCommand();
            selectCmd.Connection = conn;
            selectCmd.CommandText = "SELECT 干毛焦,湿焦粉,焦丁,焦炭水份,焦炭灰份,焦炭挥发份,焦炭S,焦炭C,焦炭M25,焦炭M10 FROM 每月焦 Where 年=:年 and 月=:月";
            selectCmd.Parameters.Add(":年", OracleType.Int32).Value = entity.Nian;
            selectCmd.Parameters.Add(":月", OracleType.Int32).Value = entity.Yue;

            OracleDataReader dr = selectCmd.ExecuteReader();
            if (dr.Read())
            {
                if (dr.IsDBNull(0)) entity.GanMaoJiao = null; else entity.GanMaoJiao = dr.GetDouble(0);
                if (dr.IsDBNull(1)) entity.ShiJiaoFen = null; else entity.ShiJiaoFen = dr.GetDouble(1);
                if (dr.IsDBNull(2)) entity.JiaoDing = null; else entity.JiaoDing = dr.GetDouble(2);
                if (dr.IsDBNull(3)) entity.JiaoShuiFen = null; else entity.JiaoShuiFen = dr.GetDouble(3);
                if (dr.IsDBNull(4)) entity.JiaoHuiFen = null; else entity.JiaoHuiFen = dr.GetDouble(4);
                if (dr.IsDBNull(5)) entity.JiaoHuiFaFen = null; else entity.JiaoHuiFaFen = dr.GetDouble(5);
                if (dr.IsDBNull(6)) entity.JiaoS = null; else entity.JiaoS = dr.GetDouble(6);
                if (dr.IsDBNull(7)) entity.JiaoC = null; else entity.JiaoC = dr.GetDouble(7);
                if (dr.IsDBNull(8)) entity.JiaoM25 = null; else entity.JiaoM25 = dr.GetDouble(8);
                if (dr.IsDBNull(9)) entity.JiaoM10 = null; else entity.JiaoM10 = dr.GetDouble(9);
            }
            dr.Close();

            OracleCommand selectCmdMei = new OracleCommand();
            selectCmdMei.Connection = conn;
            selectCmdMei.CommandText = "SELECT 重量,水份,灰份,挥发份,S FROM 每月煤 Where 年=:年 and 月=:月";
            selectCmdMei.Parameters.Add(":年", OracleType.Int32).Value = entity.Nian;
            selectCmdMei.Parameters.Add(":月", OracleType.Int32).Value = entity.Yue;

            OracleDataReader drMei = selectCmdMei.ExecuteReader();
            if (drMei.Read())
            {
                if (drMei.IsDBNull(0)) entity.ShiMeiFen = null; else entity.ShiMeiFen = drMei.GetDouble(0);
                if (drMei.IsDBNull(1)) entity.MeiShuiFen = null; else entity.MeiShuiFen = drMei.GetDouble(1);
                if (drMei.IsDBNull(2)) entity.MeiHuiFen = null; else entity.MeiHuiFen = drMei.GetDouble(2);
                if (drMei.IsDBNull(3)) entity.MeiHuiFaFen = null; else entity.MeiHuiFaFen = drMei.GetDouble(3);
                if (drMei.IsDBNull(4)) entity.MeiS = null; else entity.MeiS = drMei.GetDouble(4);

            }
            drMei.Close();

            OracleCommand selectCmdEnergy = new OracleCommand();
            selectCmdEnergy.Connection = conn;
            selectCmdEnergy.CommandText = "SELECT 耗电量,蒸汽,耗水,高炉煤气,煤气回收,富氧量,氮气,大块铁,高炉1人数,高炉2人数,高炉3人数,高炉4人数,高炉5人数,高炉6人数 FROM 每月能耗  Where 年=:年 and 月=:月";
            selectCmdEnergy.Parameters.Add(":年", OracleType.Int32).Value = entity.Nian;
            selectCmdEnergy.Parameters.Add(":月", OracleType.Int32).Value = entity.Yue;

            OracleDataReader drEnergy = selectCmdEnergy.ExecuteReader();
            if (drEnergy.Read())
            {
                if (drEnergy.IsDBNull(0)) entity.HaoDianLiang = null; else entity.HaoDianLiang = drEnergy.GetDouble(0);
                if (drEnergy.IsDBNull(1)) entity.ZhengQi = null; else entity.ZhengQi = drEnergy.GetDouble(1);
                if (drEnergy.IsDBNull(2)) entity.HaoShui = null; else entity.HaoShui = drEnergy.GetDouble(2);
                if (drEnergy.IsDBNull(3)) entity.GaoLuMeiQi = null; else entity.GaoLuMeiQi = drEnergy.GetDouble(3);
                if (drEnergy.IsDBNull(4)) entity.HuiShouMeiQi = null; else entity.HuiShouMeiQi = drEnergy.GetDouble(4);
                if (drEnergy.IsDBNull(5)) entity.FuYang = null; else entity.FuYang = drEnergy.GetDouble(5);
                if (drEnergy.IsDBNull(6)) entity.DanQi = null; else entity.DanQi = drEnergy.GetDouble(6);
                if (drEnergy.IsDBNull(7)) entity.DaKuaiTie = null; else entity.DaKuaiTie = drEnergy.GetDouble(7);
                if (drEnergy.IsDBNull(8)) entity.GaoLu1RenShu = null; else entity.GaoLu1RenShu = drEnergy.GetInt32(8);
                if (drEnergy.IsDBNull(9)) entity.GaoLu2RenShu = null; else entity.GaoLu2RenShu = drEnergy.GetInt32(9);
                if (drEnergy.IsDBNull(10)) entity.GaoLu3RenShu = null; else entity.GaoLu3RenShu = drEnergy.GetInt32(10);
                if (drEnergy.IsDBNull(11)) entity.GaoLu4RenShu = null; else entity.GaoLu4RenShu = drEnergy.GetInt32(11);
                if (drEnergy.IsDBNull(12)) entity.GaoLu5RenShu = null; else entity.GaoLu5RenShu = drEnergy.GetInt32(12);
                if (drEnergy.IsDBNull(13)) entity.GaoLu6RenShu = null; else entity.GaoLu6RenShu = drEnergy.GetInt32(13);
            }
            drEnergy.Close();


            conn.Close();
            return entity;
        }
    }
}
