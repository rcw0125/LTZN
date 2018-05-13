using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;

using LTZN.Data;

namespace LTZN.Repository
{
    public class RiLianTieXiaoHaoRepository : RepositoryBase
    {
        public RiLianTieXiaoHao GetRiLianTieXiaoHao(DateTime riqi)
        {
            conn.Open();
            RiLianTieXiaoHao result = new RiLianTieXiaoHao();
            result.RiQi = riqi;

            OracleCommand selectCmd = new OracleCommand();
            selectCmd.Connection = conn;
            selectCmd.CommandText = "SELECT P02自产焦水分, P03落地焦水分, P04焦粉水分, P05二号皮带, P06三号皮带, P07总返矿, P08烧结二号称, P09备注1, P10备注2 FROM 全厂日消耗 where P01日期=:RiQi";
            selectCmd.Parameters.Add(":RiQi", OracleType.DateTime).Value = result.RiQi;

            OracleDataReader dr = selectCmd.ExecuteReader();
            if (dr.Read())
            {
                if (dr.IsDBNull(0)) result.ZiChanJiaoShuiFen = null; else result.ZiChanJiaoShuiFen = dr.GetDouble(0);
                if (dr.IsDBNull(1)) result.LuoDiJiaoShuiFen = null; else result.LuoDiJiaoShuiFen = dr.GetDouble(1);
                if (dr.IsDBNull(2)) result.JiaoFenShuiFen = null; else result.JiaoFenShuiFen = dr.GetDouble(2);
                if (dr.IsDBNull(3)) result.ErHaoPiDai = null; else result.ErHaoPiDai = dr.GetDouble(3);
                if (dr.IsDBNull(4)) result.SanHaoPiDai = null; else result.SanHaoPiDai = dr.GetDouble(4);
                if (dr.IsDBNull(5)) result.ZongFanKuang = null; else result.ZongFanKuang = dr.GetDouble(5);
                if (dr.IsDBNull(6)) result.ShaojieErHaoCheng = null; else result.ShaojieErHaoCheng = dr.GetDouble(6);
                if (dr.IsDBNull(7)) result.BeiZhu1 = ""; else result.BeiZhu1 = dr.GetString(7);
                if (dr.IsDBNull(8)) result.BeiZhu2 = ""; else result.BeiZhu2 = dr.GetString(8);
            }

            dr.Close();
            conn.Close();
            return result;
        }

        public void Save(RiLianTieXiaoHao entity)
        {
            conn.Open();
            OracleCommand cmd1 = new OracleCommand();
            cmd1.Connection = conn;
            cmd1.CommandText = "Select count(*) from 全厂日消耗 where P01日期=:RiQi";
            cmd1.Parameters.Add(":RiQi", OracleType.Int32).Value = entity.RiQi;
            if (Convert.ToInt32(cmd1.ExecuteScalar()) == 0)
            {
                OracleCommand insertCmd = new OracleCommand();
                insertCmd.Connection = conn;
                insertCmd.CommandText = "INSERT INTO 全厂日消耗(P01日期, P02自产焦水分, P03落地焦水分, P04焦粉水分, P05二号皮带, P06三号皮带, P07总返矿, P08烧结二号称, P09备注1, P10备注2) VALUES (:P01日期, :P02自产焦水分, :P03落地焦水分, :P04焦粉水分, :P05二号皮带, :P06三号皮带, :P07总返矿, :P08烧结二号称, :P09备注1, :P10备注2)";

                insertCmd.Parameters.Add(":P01日期", OracleType.DateTime).Value = entity.RiQi;
                insertCmd.Parameters.Add(":P02自产焦水分", OracleType.Double, 22).Value = (object)entity.ZiChanJiaoShuiFen ?? DBNull.Value;
                insertCmd.Parameters.Add(":P03落地焦水分", OracleType.Double, 22).Value = (object)entity.LuoDiJiaoShuiFen ?? DBNull.Value;
                insertCmd.Parameters.Add(":P04焦粉水分", OracleType.Double, 22).Value = (object)entity.JiaoFenShuiFen ?? DBNull.Value;
                insertCmd.Parameters.Add(":P05二号皮带", OracleType.Double, 22).Value = (object)entity.ErHaoPiDai ?? DBNull.Value;
                insertCmd.Parameters.Add(":P06三号皮带", OracleType.Double, 22).Value = (object)entity.SanHaoPiDai ?? DBNull.Value;
                insertCmd.Parameters.Add(":P07总返矿", OracleType.Double, 22).Value = (object)entity.ZongFanKuang ?? DBNull.Value;
                insertCmd.Parameters.Add(":P08烧结二号称", OracleType.Double, 22).Value = (object)entity.ShaojieErHaoCheng ?? DBNull.Value;
                insertCmd.Parameters.Add(":P09备注1", OracleType.VarChar).Value = (object)entity.BeiZhu1 ?? DBNull.Value;
                insertCmd.Parameters.Add(":P10备注2", OracleType.VarChar).Value = (object)entity.BeiZhu2 ?? DBNull.Value;

                insertCmd.ExecuteNonQuery();

            }
            else
            {

                OracleCommand updateCmd = new OracleCommand();
                updateCmd.Connection = conn;
                updateCmd.CommandText = "UPDATE 全厂日消耗 SET P01日期 = :P01日期, P02自产焦水分 = :P02自产焦水分, P03落地焦水分 = :P03落地焦水分, P04焦粉水分 = :P04焦粉水分, P05二号皮带 = :P05二号皮带, P06三号皮带 = :P06三号皮带, P07总返矿 = :P07总返矿, P08烧结二号称 = :P08烧结二号称, P09备注1 = :P09备注1, P10备注2 = :P10备注2 WHERE P01日期 = :P01日期";


                updateCmd.Parameters.Add(":P02自产焦水分", OracleType.Double, 22).Value = (object)entity.ZiChanJiaoShuiFen ?? DBNull.Value;
                updateCmd.Parameters.Add(":P03落地焦水分", OracleType.Double, 22).Value = (object)entity.LuoDiJiaoShuiFen ?? DBNull.Value;
                updateCmd.Parameters.Add(":P04焦粉水分", OracleType.Double, 22).Value = (object)entity.JiaoFenShuiFen ?? DBNull.Value;
                updateCmd.Parameters.Add(":P05二号皮带", OracleType.Double, 22).Value = (object)entity.ErHaoPiDai ?? DBNull.Value;
                updateCmd.Parameters.Add(":P06三号皮带", OracleType.Double, 22).Value = (object)entity.SanHaoPiDai ?? DBNull.Value;
                updateCmd.Parameters.Add(":P07总返矿", OracleType.Double, 22).Value = (object)entity.ZongFanKuang ?? DBNull.Value;
                updateCmd.Parameters.Add(":P08烧结二号称", OracleType.Double, 22).Value = (object)entity.ShaojieErHaoCheng ?? DBNull.Value;
                updateCmd.Parameters.Add(":P09备注1", OracleType.VarChar).Value = (object)entity.BeiZhu1 ?? DBNull.Value;
                updateCmd.Parameters.Add(":P10备注2", OracleType.VarChar).Value = (object)entity.BeiZhu2 ?? DBNull.Value;
                updateCmd.Parameters.Add(":P01日期", OracleType.DateTime).Value = entity.RiQi;

                updateCmd.ExecuteNonQuery();
            }
            conn.Close();

        }
    }
}
