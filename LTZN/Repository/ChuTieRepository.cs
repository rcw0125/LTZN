using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;
using LTZN.Data;
namespace LTZN.Repository
{
    public class ChuTieRepository : RepositoryBase
    {
        private ChuTieJiHuaRepository _chuTieJiHuaRepository = null;

        public ChuTieJiHuaRepository ChuTieJiHuaRepository
        {
            get
            {
                if (_chuTieJiHuaRepository == null)
                {
                    _chuTieJiHuaRepository = new ChuTieJiHuaRepository();
                    _chuTieJiHuaRepository.LoadData();
                }
                return _chuTieJiHuaRepository;
            }
        }
       
        public ChuTie GetChuTie(int gaolu, DateTime riqi, string banci, int banxuhao)
        {
            DateTime? zdsj = ChuTieJiHuaRepository.GetZdsj(gaolu, riqi, banci, banxuhao);
            if (zdsj.HasValue)
                return GetChuTie(gaolu, zdsj.Value);
            else
                throw new Exception("无法找到出铁计划数据。");

        }

        public ChuTie GetChuTie(int gaolu, DateTime zdsj)
        {
            ChuTie result = null;
            conn.Open();
            OracleCommand cmdQuery = new OracleCommand();
            cmdQuery.Connection = conn;
            cmdQuery.CommandText = "Select LUCI,DGSJ,DKSJ,TZSJ,BANCI,BANLUCI,WDSJ,QUCHU,FELIANG,FEC,FESI,FEMN,FEP,FES,FETI,ZHASIO2,ZHACAO,ZHAMGO,ZHAAL2O3,ZHAS,ZHATIO2,ZHAR2,ZHAR3,ZHAR4,ZHAMGOALO,LIAOPI from DDLUCI Where GAOLU=:gaolu and ZDSJ=:zdsj";
            cmdQuery.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            cmdQuery.Parameters.Add(":zdsj", OracleType.DateTime).Value = zdsj;
            OracleDataReader dr = cmdQuery.ExecuteReader();
            while (dr.Read())
            {
                result = new ChuTie();
                result.GaoLu = gaolu;
                result.Zdsj = zdsj;
                result.LuCiHao = dr.IsDBNull(0) ? "" : dr.GetString(0);
                result.Dgsj = dr.IsDBNull(1) ? null : (DateTime?)dr.GetDateTime(1);
                result.Dksj = dr.IsDBNull(2) ? null : (DateTime?)dr.GetDateTime(2);
                result.Tzsj = dr.IsDBNull(3) ? null : (DateTime?)dr.GetDateTime(3);
                result.BanCi = dr.IsDBNull(4) ? "" : dr.GetString(4);
                result.BanXuHao = dr.IsDBNull(5) ? 0 : dr.GetInt32(5);
                result.Wdsj = dr.IsDBNull(6) ? null : (int?)dr.GetInt32(6);
                result.QuChu = dr.IsDBNull(7) ? "炼钢" : dr.GetString(7);
                result.FeC = dr.IsDBNull(8) ? null : (double?)dr.GetDouble(8);
                result.FeSi = dr.IsDBNull(9) ? null : (double?)dr.GetDouble(9);
                result.FeMn = dr.IsDBNull(10) ? null : (double?)dr.GetDouble(10);
                result.FeP = dr.IsDBNull(11) ? null : (double?)dr.GetDouble(11);
                result.FeS = dr.IsDBNull(12) ? null : (double?)dr.GetDouble(12);
                result.FeTi = dr.IsDBNull(13) ? null : (double?)dr.GetDouble(13);
                result.ZhaSiO2 = dr.IsDBNull(14) ? null : (double?)dr.GetDouble(14);
                result.ZhaCaO = dr.IsDBNull(15) ? null : (double?)dr.GetDouble(15);
                result.ZhaMgO = dr.IsDBNull(16) ? null : (double?)dr.GetDouble(16);
                result.ZhaAl2O3 = dr.IsDBNull(17) ? null : (double?)dr.GetDouble(17);
                result.ZhaS = dr.IsDBNull(18) ? null : (double?)dr.GetDouble(18);
                result.ZhaTiO2 = dr.IsDBNull(19) ? null : (double?)dr.GetDouble(19);
                result.ZhaR2 = dr.IsDBNull(20) ? null : (double?)dr.GetDouble(20);
                result.ZhaR3 = dr.IsDBNull(21) ? null : (double?)dr.GetDouble(21);
                result.ZhaR4 = dr.IsDBNull(22) ? null : (double?)dr.GetDouble(22);
                result.ZhaMgOAlO = dr.IsDBNull(23) ? null : (double?)dr.GetDouble(23);
                result.LiaoPiShu = dr.IsDBNull(24) ? null : (int?)dr.GetInt32(24);
                result.ReSet();
            }
            dr.Close();
            conn.Close();
            return result;
        }

        /// <summary>
        /// 得到最后的堵口时间
        /// </summary>
        /// <param name="gaolu"></param>
        /// <returns></returns>
        public DateTime? GetMaxDksj(int gaolu)
        {
            DateTime? result = null;
            conn.Open();
            OracleCommand cmdQuery = new OracleCommand();
            cmdQuery.Connection = conn;
            cmdQuery.CommandText = "Select Max(DKSJ) from DDLUCI Where GAOLU=:gaolu and DKSJ is not null";
            cmdQuery.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = cmdQuery.ExecuteReader();
            if (dr.Read())
            {
                if (!dr.IsDBNull(0))
                    result = dr.GetDateTime(0);
            }
            dr.Close();
            conn.Close();
            return result;
        }

        /// <summary>
        /// 得到正点时间前的最后的堵口时间
        /// </summary>
        /// <param name="gaolu"></param>
        /// <returns></returns>
        public DateTime? GetMaxDksj(int gaolu, DateTime zdsj)
        {
            DateTime? result = null;
            conn.Open();
            OracleCommand cmdQuery = new OracleCommand();
            cmdQuery.Connection = conn;
            cmdQuery.CommandText = "Select Max(DKSJ) from DDLUCI Where GAOLU=:gaolu and ZDSJ<:zdsj and DKSJ is not null";
            cmdQuery.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            cmdQuery.Parameters.Add(":zdsj", OracleType.DateTime).Value = zdsj;
            OracleDataReader dr = cmdQuery.ExecuteReader();
            if (dr.Read())
            {
                if (!dr.IsDBNull(0))
                    result = dr.GetDateTime(0);
            }
            dr.Close();
            conn.Close();
            return result;
        }

        public int? GetLiaoPiShu(int gaolu, DateTime dksj1, DateTime dksj2)
        {
            int? result = null;
            conn.Open();
            OracleCommand cmdQuery = new OracleCommand();
            cmdQuery.Connection = conn;
            cmdQuery.CommandText = "select count(distinct pishu) into lp from lt_liao where GAOLU=:gaolu and t>:dksj1 and t<=:dksj2";
            cmdQuery.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            cmdQuery.Parameters.Add(":dksj1", OracleType.DateTime).Value = dksj1;
            cmdQuery.Parameters.Add(":dksj2", OracleType.DateTime).Value = dksj2;
            OracleDataReader dr = cmdQuery.ExecuteReader();
            if (dr.Read())
            {
                if (!dr.IsDBNull(0))
                    result = dr.GetInt32(0);
            }
            dr.Close();
            conn.Close();
            return result;
        }

        public void SaveChuTie(ChuTie chutie)
        {
            if (chutie.EntityState == EntityState.Unmodified) return;
            conn.Open();
            OracleCommand cmd1 = new OracleCommand();
            cmd1.Connection = conn;
            cmd1.CommandText = "Select count(*) from DDLUCI Where GAOLU=:gaolu and ZDSJ=:zdsj";
            cmd1.Parameters.Add(":gaolu", OracleType.Int32).Value = chutie.GaoLu;
            cmd1.Parameters.Add(":zdsj", OracleType.DateTime).Value = chutie.Zdsj;
            if (Convert.ToInt32(cmd1.ExecuteScalar()) == 0)
            {
                OracleCommand insertCmd = new OracleCommand();
                insertCmd.Connection = conn;
                StringBuilder sbInsertSql = new StringBuilder();
                sbInsertSql.Append("INSERT INTO DDLUCI(");
                sbInsertSql.Append("GAOLU,ZDSJ,LUCI,DGSJ,DKSJ,TZSJ,BANCI,BANLUCI,WDSJ,QUCHU,FELIANG,FEC,FESI,FEMN,FEP,FES,FETI,ZHASIO2,ZHACAO,ZHAMGO,ZHAAL2O3,ZHAS,ZHATIO2,ZHAR2,ZHAR3,ZHAR4,ZHAMGOALO,LIAOPI");
                sbInsertSql.Append(") VALUES (");
                sbInsertSql.Append(":GAOLU,:ZDSJ,:LUCI,:DGSJ,:DKSJ,:TZSJ,:BANCI,:BANLUCI,:WDSJ,:QUCHU,:FELIANG,:FEC,:FESI,:FEMN,:FEP,:FES,:FETI,:ZHASIO2,:ZHACAO,:ZHAMGO,:ZHAAL2O3,:ZHAS,:ZHATIO2,:ZHAR2,GETR3(:ZHACAO,:ZHAMGO,:ZHASIO2),GETR4(:ZHACAO,:ZHAMGO,:ZHASIO2,:ZHAAL2O3),GETMGOALO(:ZHAMGO,:ZHAAL2O3),:LIAOPI");
                sbInsertSql.Append(")");
                insertCmd.CommandText = sbInsertSql.ToString();

                insertCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = chutie.GaoLu;
                insertCmd.Parameters.Add(":ZDSJ", OracleType.DateTime).Value = chutie.Zdsj;
                insertCmd.Parameters.Add(":LUCI", OracleType.VarChar, 20).Value = chutie.LuCiHao ?? "";
                insertCmd.Parameters.Add(":DGSJ", OracleType.DateTime).Value = (object)chutie.Dgsj ?? DBNull.Value;
                insertCmd.Parameters.Add(":DKSJ", OracleType.DateTime).Value = (object)chutie.Dksj ?? DBNull.Value;
                insertCmd.Parameters.Add(":TZSJ", OracleType.DateTime).Value = (object)chutie.Tzsj ?? DBNull.Value;
                insertCmd.Parameters.Add(":BANCI", OracleType.VarChar, 20).Value = chutie.BanCi ?? "";
                insertCmd.Parameters.Add(":BANLUCI", OracleType.Int32).Value = chutie.BanXuHao;
                insertCmd.Parameters.Add(":WDSJ", OracleType.Int32).Value = (object)chutie.Wdsj ?? DBNull.Value;
                insertCmd.Parameters.Add(":QUCHU", OracleType.VarChar, 20).Value = chutie.QuChu ?? "";
                insertCmd.Parameters.Add(":FELIANG", OracleType.Double, 22).Value = (object)chutie.ChanLiang ?? DBNull.Value;
                insertCmd.Parameters.Add(":FEC", OracleType.Double, 22).Value = (object)chutie.FeC ?? DBNull.Value;
                insertCmd.Parameters.Add(":FESI", OracleType.Double, 22).Value = (object)chutie.FeSi ?? DBNull.Value;
                insertCmd.Parameters.Add(":FEMN", OracleType.Double, 22).Value = (object)chutie.FeMn ?? DBNull.Value;
                insertCmd.Parameters.Add(":FEP", OracleType.Double, 22).Value = (object)chutie.FeP ?? DBNull.Value;
                insertCmd.Parameters.Add(":FES", OracleType.Double, 22).Value = (object)chutie.FeS ?? DBNull.Value;
                insertCmd.Parameters.Add(":FETI", OracleType.Double, 22).Value = (object)chutie.FeTi ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHASIO2", OracleType.Double, 22).Value = (object)chutie.ZhaSiO2 ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHACAO", OracleType.Double, 22).Value = (object)chutie.ZhaCaO ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHAMGO", OracleType.Double, 22).Value = (object)chutie.ZhaMgO ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHAAL2O3", OracleType.Double, 22).Value = (object)chutie.ZhaAl2O3 ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHAS", OracleType.Double, 22).Value = (object)chutie.ZhaS ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHATIO2", OracleType.Double, 22).Value = (object)chutie.ZhaTiO2 ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHAR2", OracleType.Double, 22).Value = (object)chutie.ZhaR2 ?? DBNull.Value;
                // insertCmd.Parameters.Add(":ZHAR3", OracleType.Double, 22).Value = (object)chutie.ZhaR3 ?? DBNull.Value;
                // insertCmd.Parameters.Add(":ZHAR4", OracleType.Double, 22).Value = (object)chutie.ZhaR4 ?? DBNull.Value;
                //insertCmd.Parameters.Add(":ZHAMGOALO", OracleType.Double, 22).Value = (object)chutie.ZhaMgOAlO ?? DBNull.Value;
                insertCmd.Parameters.Add(":LIAOPI", OracleType.Int32).Value = (object)chutie.LiaoPiShu ?? DBNull.Value;
                insertCmd.ExecuteNonQuery();

            }
            else
            {
                OracleCommand updateCmd = new OracleCommand();
                updateCmd.Connection = conn;
                StringBuilder sbUpdateSql = new StringBuilder();
                sbUpdateSql.Append("UPDATE DDLUCI ");
                sbUpdateSql.Append("SET   DGSJ =:DGSJ, DKSJ =:DKSJ,TZSJ =:TZSJ,BANCI=:BANCI,BANLUCI=:BANLUCI,WDSJ=:WDSJ,QUCHU =:QUCHU,FELIANG =:FELIANG,");
                sbUpdateSql.Append(" FEC =:FEC,FESI =:FESI, FEMN =:FEMN, FEP =:FEP, FES =:FES, FETI =:FETI,");
                sbUpdateSql.Append(" ZHASIO2 =:ZHASIO2, ZHACAO =:ZHACAO, ZHAMGO =:ZHAMGO,ZHAAL2O3 =:ZHAAL2O3, ZHAS =:ZHAS,ZHATIO2=:ZHATIO2,");
                sbUpdateSql.Append(" ZHAR2=:ZHAR2,LIAOPI=:LIAOPI,ZHAR3=GETR3(ZHACAO,ZHAMGO,ZHASIO2),ZHAR4=GETR4(ZHACAO,ZHAMGO,ZHASIO2,ZHAAL2O3),ZHAMGOALO=GETMGOALO(ZHAMGO,ZHAAL2O3)");
                sbUpdateSql.Append(" Where GAOLU=:GAOLU and ZDSJ=:ZDSJ");

                updateCmd.CommandText = sbUpdateSql.ToString();

                updateCmd.Parameters.Add(":LUCI", OracleType.VarChar, 20).Value = chutie.LuCiHao ?? "";
                updateCmd.Parameters.Add(":DGSJ", OracleType.DateTime).Value = (object)chutie.Dgsj ?? DBNull.Value;
                updateCmd.Parameters.Add(":DKSJ", OracleType.DateTime).Value = (object)chutie.Dksj ?? DBNull.Value;
                updateCmd.Parameters.Add(":TZSJ", OracleType.DateTime).Value = (object)chutie.Tzsj ?? DBNull.Value;
                updateCmd.Parameters.Add(":BANCI", OracleType.VarChar, 20).Value = chutie.BanCi ?? "";
                updateCmd.Parameters.Add(":BANLUCI", OracleType.Int32).Value = chutie.BanXuHao;
                updateCmd.Parameters.Add(":WDSJ", OracleType.Int32).Value = (object)chutie.Wdsj ?? DBNull.Value;
                updateCmd.Parameters.Add(":QUCHU", OracleType.VarChar, 20).Value = chutie.QuChu ?? "";
                updateCmd.Parameters.Add(":FELIANG", OracleType.Double, 22).Value = (object)chutie.ChanLiang ?? DBNull.Value;
                updateCmd.Parameters.Add(":FEC", OracleType.Double, 22).Value = (object)chutie.FeC ?? DBNull.Value;
                updateCmd.Parameters.Add(":FESI", OracleType.Double, 22).Value = (object)chutie.FeSi ?? DBNull.Value;
                updateCmd.Parameters.Add(":FEMN", OracleType.Double, 22).Value = (object)chutie.FeMn ?? DBNull.Value;
                updateCmd.Parameters.Add(":FEP", OracleType.Double, 22).Value = (object)chutie.FeP ?? DBNull.Value;
                updateCmd.Parameters.Add(":FES", OracleType.Double, 22).Value = (object)chutie.FeS ?? DBNull.Value;
                updateCmd.Parameters.Add(":FETI", OracleType.Double, 22).Value = (object)chutie.FeTi ?? DBNull.Value;
                updateCmd.Parameters.Add(":ZHASIO2", OracleType.Double, 22).Value = (object)chutie.ZhaSiO2 ?? DBNull.Value;
                updateCmd.Parameters.Add(":ZHACAO", OracleType.Double, 22).Value = (object)chutie.ZhaCaO ?? DBNull.Value;
                updateCmd.Parameters.Add(":ZHAMGO", OracleType.Double, 22).Value = (object)chutie.ZhaMgO ?? DBNull.Value;
                updateCmd.Parameters.Add(":ZHAAL2O3", OracleType.Double, 22).Value = (object)chutie.ZhaAl2O3 ?? DBNull.Value;
                updateCmd.Parameters.Add(":ZHAS", OracleType.Double, 22).Value = (object)chutie.ZhaS ?? DBNull.Value;
                updateCmd.Parameters.Add(":ZHATIO2", OracleType.Double, 22).Value = (object)chutie.ZhaTiO2 ?? DBNull.Value;
                updateCmd.Parameters.Add(":ZHAR2", OracleType.Double, 22).Value = (object)chutie.ZhaR2 ?? DBNull.Value;

                // updateCmd.Parameters.Add(":ZHAR3", OracleType.Double, 22).Value = (object)chutie.ZhaR3 ?? DBNull.Value;
                // updateCmd.Parameters.Add(":ZHAR4", OracleType.Double, 22).Value = (object)chutie.ZhaR4 ?? DBNull.Value;    
                //updateCmd.Parameters.Add(":ZHAMGOALO", OracleType.Double, 22).Value = (object)chutie.ZhaMgOAlO ?? DBNull.Value;
                updateCmd.Parameters.Add(":LIAOPI", OracleType.Int32).Value = (object)chutie.LiaoPiShu ?? DBNull.Value;
                updateCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = chutie.GaoLu;
                updateCmd.Parameters.Add(":ZDSJ", OracleType.DateTime).Value = chutie.Zdsj;
                updateCmd.ExecuteNonQuery();
            }
            conn.Close();
        }
    }
}
