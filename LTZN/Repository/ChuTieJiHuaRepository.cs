using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;
using LTZN.Data;

namespace LTZN.Repository
{
    public class ChuTieJiHuaRepository:RepositoryBase
    {
        private List<ChuTieJiHua> jiHuaData = new List<ChuTieJiHua>();

        public List<ChuTieJiHua> JiHuaData
        {
            get { return jiHuaData; }
        }

        public void LoadData()
        {
            conn.Open();
            OracleCommand cmdQuery = new OracleCommand();
            cmdQuery.Connection = conn;
            cmdQuery.CommandText = "Select GAOLU,XUHAO,ZDSJ,OFFSET,BANCI,BANXUHAO from CHUTIE_JIHUA";
            OracleDataReader dr = cmdQuery.ExecuteReader();
            while (dr.Read())
            {
                ChuTieJiHua jihua = new ChuTieJiHua();
                jihua.GaoLu = dr.GetInt32(0);
                jihua.XuHao = dr.GetInt32(1);
                jihua.Zdsj = dr.GetInt32(2);
                jihua.OffSet = dr.GetInt32(3);
                jihua.BanCi = dr.GetString(4);
                jihua.BanXuHao = dr.GetInt32(5);
                jihua.ReSet();
                jiHuaData.Add(jihua);
            }
            dr.Close();
            conn.Close();
        }

        public void Save(ChuTieJiHua jihua)
        {
            if (jihua.EntityState == EntityState.Unmodified) return;

            conn.Open();
            OracleCommand cmd1 = new OracleCommand();
            cmd1.Connection = conn;
            cmd1.CommandText = "Select count(*) from CHUTIE_JIHUA Where GAOLU=:gaolu and XUHAO=:xuhao";
            cmd1.Parameters.Add(":gaolu", OracleType.Int32).Value = jihua.GaoLu;
            cmd1.Parameters.Add(":xuhao", OracleType.Int32).Value = jihua.XuHao;
            if (Convert.ToInt32(cmd1.ExecuteScalar()) == 0)
            {
                OracleCommand insertCmd = new OracleCommand();
                insertCmd.Connection = conn;
                insertCmd.CommandText = "INSERT INTO CHUTIE_JIHUA(GAOLU,XUHAO,ZDSJ,OFFSET,BANCI,BANXUHAO) VALUES(:gaolu,:xuhao,:zdsj,:offset,:banci,:banxuhao)";
                insertCmd.Parameters.Add(":gaolu", OracleType.Int32).Value = jihua.GaoLu;
                insertCmd.Parameters.Add(":xuhao", OracleType.Int32).Value = jihua.XuHao;
                insertCmd.Parameters.Add(":zdsj", OracleType.Int32).Value = jihua.Zdsj;
                insertCmd.Parameters.Add(":offset", OracleType.Int32).Value = jihua.OffSet;
                insertCmd.Parameters.Add(":banci", OracleType.VarChar, 200).Value = jihua.BanCi;
                insertCmd.Parameters.Add(":banxuhao", OracleType.Int32).Value = jihua.BanXuHao;
                insertCmd.ExecuteNonQuery();
            }
            else
            {
                OracleCommand updateCmd = new OracleCommand();
                updateCmd.Connection = conn;
                updateCmd.CommandText = "UPDATE CHUTIE_JIHUA SET ZDSJ=:zdsj,OFFSET=:offset,BANCI=:banci,BANXUHAO=:banxuhao WHERE GAOLU=:gaolu and XUHAO=:xuhao";

                updateCmd.Parameters.Add(":zdsj", OracleType.Int32).Value = jihua.Zdsj;
                updateCmd.Parameters.Add(":offset", OracleType.Int32).Value = jihua.OffSet;
                updateCmd.Parameters.Add(":banci", OracleType.VarChar, 200).Value = jihua.BanCi;
                updateCmd.Parameters.Add(":banxuhao", OracleType.Int32).Value = jihua.BanXuHao;
                
                updateCmd.Parameters.Add(":gaolu", OracleType.Int32).Value = jihua.GaoLu;
                updateCmd.Parameters.Add(":xuhao", OracleType.Int32).Value = jihua.XuHao;
                updateCmd.ExecuteNonQuery();
            }
            conn.Close();

        }

        public void Delete(ChuTieJiHua jihua)
        {
            conn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;
            cmd.CommandText = "Delete From CHUTIE_JIHUA WHERE GAOLU=:gaolu and XUHAO=:xuhao";
            cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = jihua.GaoLu;
            cmd.Parameters.Add(":xuhao", OracleType.Int32).Value = jihua.XuHao;
            cmd.ExecuteNonQuery();
            conn.Close();
        }

        /// <summary>
        /// 获取正点时间
        /// </summary>
        /// <param name="gaolu"></param>
        /// <param name="riqi"></param>
        /// <param name="banci"></param>
        /// <param name="banxuhao"></param>
        /// <returns></returns>
        public DateTime? GetZdsj(int gaolu, DateTime riqi, string banci, int banxuhao)
        {
            DateTime? result = null;
            foreach (var item in jiHuaData)
            {
                if (item.GaoLu == gaolu && item.BanCi == banci && item.BanXuHao == banxuhao)
                {
                    result = riqi.AddMinutes(item.Zdsj - item.OffSet);
                    break;
                }
            }
            return result;
        }

        public List<DateTime> GetZdsjToNow(int gaolu, DateTime lastZdsj)
        {
            List<DateTime> result = new List<DateTime>();
            DateTime beginDate = lastZdsj.Date;
            while (beginDate <= DateTime.Today.AddDays(1))
            {
                foreach (var item in jiHuaData)
                {
                    DateTime zdsj = beginDate.AddMinutes(item.Zdsj - item.OffSet);
                    if (zdsj > lastZdsj && zdsj < DateTime.Now.AddMonths(12))
                    {
                        result.Add(zdsj);
                    }
                }
                beginDate = beginDate.AddDays(1);
            }
            return result;
        }

    }
}
