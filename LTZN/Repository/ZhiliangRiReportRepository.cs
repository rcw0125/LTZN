using System;
using System.Collections.Generic;
using System.Data.OracleClient;
using System.Text;
using LTZN.质量日报;
namespace LTZN.Repository
{
    class ZhiliangRiReportRepository : RepositoryBase
    {
        public 质量日报内容 GetRiReport(DateTime riqi)
        {

            质量日报内容 result = new 质量日报内容(riqi);

            StringBuilder sbSql = new StringBuilder();

            sbSql.AppendLine("SELECT MC,round(avg(TFE),3),round(avg(SIO),3),round(avg(CAO),3),round(avg(FEO),3),round(avg(MGO),3),round(avg(S),3),round(avg(ALO),3),round(avg(TIO2),3),round(avg(P),3),round(avg(MNO),3),round(avg(CR),3),round(avg(PB),3),round(avg(ZN),3),round(avg(JJS),3),round(avg(V2O5),3),round(avg(CU),3),round(avg(MO),3),round(avg(NI),3),round(avg(SN),3),round(avg(SB),3),round(avg(AS1),3),round(avg(JD),3) FROM DDYL WHERE trunc(SJ) =to_date(:RQ) and mc not like '%高炉机烧%' group by mc");
    //        sbSql.AppendLine("FROM DDYL WHERE trunc(SJ) = :RQ)");



            conn.Open();
            OracleCommand cmduery = new OracleCommand(sbSql.ToString(), conn);
            cmduery.Parameters.Add(":RQ", OracleType.DateTime).Value = riqi;
            OracleDataReader dr = cmduery.ExecuteReader();
            while (dr.Read())

            {
                质量日报内容项 x = new 质量日报内容项(dr.GetString(0));

                x.TFE = dr.IsDBNull(1) ? null : (double?)dr.GetDouble(1);
                x.SiO2 = dr.IsDBNull(2) ? null : (double?)dr.GetDouble(2);
                x.CaO = dr.IsDBNull(3) ? null : (double?)dr.GetDouble(3);
                x.FeO = dr.IsDBNull(4) ? null : (double?)dr.GetDouble(4);
                x.MgO = dr.IsDBNull(5) ? null : (double?)dr.GetDouble(5);
                x.S = dr.IsDBNull(6) ? null : (double?)dr.GetDouble(6);
                x.Al2O3 = dr.IsDBNull(7) ? null : (double?)dr.GetDouble(7);
                x.TiO2 = dr.IsDBNull(8) ? null : (double?)dr.GetDouble(8);
                x.P = dr.IsDBNull(9) ? null : (double?)dr.GetDouble(9);
                x.MnO = dr.IsDBNull(10) ? null : (double?)dr.GetDouble(10);
                x.Cr = dr.IsDBNull(11) ? null : (double?)dr.GetDouble(11);
                x.Pb = dr.IsDBNull(12) ? null : (double?)dr.GetDouble(12);
                x.Zn = dr.IsDBNull(13) ? null : (double?)dr.GetDouble(13);
                x.JJS = dr.IsDBNull(14) ? null : (double?)dr.GetDouble(14);
                x.V2O5 = dr.IsDBNull(15) ? null : (double?)dr.GetDouble(15);
                x.Cu = dr.IsDBNull(16) ? null : (double?)dr.GetDouble(16);
                x.Mo = dr.IsDBNull(17) ? null : (double?)dr.GetDouble(17);
                x.Ni = dr.IsDBNull(18) ? null : (double?)dr.GetDouble(18);
                x.Sn = dr.IsDBNull(19) ? null : (double?)dr.GetDouble(19);
                x.Sb = dr.IsDBNull(20) ? null : (double?)dr.GetDouble(20);
                x.As = dr.IsDBNull(21) ? null : (double?)dr.GetDouble(21);
                x.R = dr.IsDBNull(22) ? null : (double?)dr.GetDouble(22);
                result.Add(x);
            }
            dr.Close();
            conn.Close();
            return result;
        
        
        
        
        
        }


        public void Save(质量日报内容 n)
        {
            conn.Open();


            conn.Close();

        }

        public 机烧粒度内容 GetJSReport(DateTime riqi)
        {
            机烧粒度内容 result = new 机烧粒度内容(riqi);

            StringBuilder sbSql = new StringBuilder();

            sbSql.AppendLine("SELECT MC,SJ,ZG,KM,SF,KY,DY40,ZJ425,ZJ2516,ZJ1610,ZJ105,XY5,RI,TBS,TBE,TB,RDI63,RDI315,RDI05 FROM DDJS WHERE trunc(SJ) =to_date(:RQ) order by SJ");

            conn.Open();
            OracleCommand cmduery = new OracleCommand(sbSql.ToString(), conn);
            cmduery.Parameters.Add(":RQ", OracleType.DateTime).Value = riqi;
            OracleDataReader dr = cmduery.ExecuteReader();
            while (dr.Read())
            {
                机烧粒度内容项 x = new 机烧粒度内容项(dr.GetString(0), dr.GetDateTime(1));

                x.ZG = dr.IsDBNull(2) ? null : (double?)dr.GetDouble(2);
                x.KM = dr.IsDBNull(3) ? null : (double?)dr.GetDouble(3);
                x.SF = dr.IsDBNull(4) ? null : (double?)dr.GetDouble(4);
                x.KY = dr.IsDBNull(5) ? null : (double?)dr.GetDouble(5);
                x.DY40 = dr.IsDBNull(6) ? null : (double?)dr.GetDouble(6);
                x.ZJ425 = dr.IsDBNull(7) ? null : (double?)dr.GetDouble(7);
                x.ZJ2516 = dr.IsDBNull(8) ? null : (double?)dr.GetDouble(8);
                x.ZJ1610 = dr.IsDBNull(9) ? null : (double?)dr.GetDouble(9);
                x.ZJ105 = dr.IsDBNull(10) ? null : (double?)dr.GetDouble(10);
                x.XY5 = dr.IsDBNull(11) ? null : (double?)dr.GetDouble(11);
                x.RI = dr.IsDBNull(12) ? null : (double?)dr.GetDouble(12);
                x.TBS = dr.IsDBNull(13) ? null : (double?)dr.GetDouble(13);
                x.TBE = dr.IsDBNull(14) ? null : (double?)dr.GetDouble(14);
                x.TB = dr.IsDBNull(15) ? null : (double?)dr.GetDouble(15);
                x.RDI63 = dr.IsDBNull(16) ? null : (double?)dr.GetDouble(16);
                x.RDI315 = dr.IsDBNull(17) ? null : (double?)dr.GetDouble(17);
                x.RDI05 = dr.IsDBNull(18) ? null : (double?)dr.GetDouble(18);
                result.Add(x);
            }
            dr.Close();
            conn.Close();
            return result; 
        
        }

        public 燃料指标内容 GetRLReport(DateTime riqi)
    {

        燃料指标内容 result = new 燃料指标内容(riqi);

        StringBuilder sbSql1 = new StringBuilder();
        sbSql1.AppendLine("SELECT round(avg(C),3),round(avg(HUIFEN),3),round(avg(HUIFA),3),round(avg(S),3),round(avg(SF),3) FROM DDMF WHERE trunc(SJ) =to_date(:RQ)");
        conn.Open();
        OracleCommand cmduery1 = new OracleCommand(sbSql1.ToString(), conn);
        cmduery1.Parameters.Add(":RQ", OracleType.DateTime).Value = riqi;
        OracleDataReader dr1 = cmduery1.ExecuteReader();
        while (dr1.Read())
        {
            燃料指标内容项 x = new 燃料指标内容项();
            x.MC = "煤粉";
            x.C = dr1.IsDBNull(0) ? null : (double?)dr1.GetDouble(0);
            x.HUIFEN = dr1.IsDBNull(1) ? null : (double?)dr1.GetDouble(1);
            x.HUIFA = dr1.IsDBNull(2) ? null : (double?)dr1.GetDouble(2);
            x.S = dr1.IsDBNull(3) ? null : (double?)dr1.GetDouble(3);
            x.SF = dr1.IsDBNull(4) ? null : (double?)dr1.GetDouble(4);
            result.Add(x);
        }
        dr1.Close();



        StringBuilder sbSql2 = new StringBuilder();
        sbSql2.AppendLine("SELECT round(avg(C),3),round(avg(HUIFEN),3),round(avg(HUIFA),3),round(avg(S),3),round(avg(SHUIF),3),round(avg(M25),3),round(avg(M10),3),round(avg(QD),3) FROM DDJT WHERE trunc(SJ) =to_date(:RQ) and mc like '自产%'");
        OracleCommand cmduery2 = new OracleCommand(sbSql2.ToString(), conn);
        cmduery2.Parameters.Add(":RQ", OracleType.DateTime).Value = riqi;
        OracleDataReader dr2 = cmduery2.ExecuteReader();
        燃料指标内容项 ZCJT = new 燃料指标内容项();
            while (dr2.Read())
        {
            ZCJT.MC = "自产焦炭";

            ZCJT.C = dr2.IsDBNull(0) ? null : (double?)dr2.GetDouble(0);
            ZCJT.HUIFEN = dr2.IsDBNull(1) ? null : (double?)dr2.GetDouble(1);
            ZCJT.HUIFA = dr2.IsDBNull(2) ? null : (double?)dr2.GetDouble(2);
            ZCJT.S = dr2.IsDBNull(3) ? null : (double?)dr2.GetDouble(3);
            ZCJT.SF = dr2.IsDBNull(4) ? null : (double?)dr2.GetDouble(4);
            ZCJT.M40 = dr2.IsDBNull(5) ? null : (double?)dr2.GetDouble(5);
            ZCJT.M10 = dr2.IsDBNull(6) ? null : (double?)dr2.GetDouble(6);
            ZCJT.QD = dr2.IsDBNull(7) ? null : (double?)dr2.GetDouble(7);
        }
            dr2.Close();
            StringBuilder sbSql3 = new StringBuilder();
            sbSql3.AppendLine("select round(avg(DY80),3),round(avg(ZJ86),3),round(avg(ZJ64),3),round(avg(ZJ42),3),round(avg(XY25),3),round(avg(PJLD),3),round(avg(LDJYX),3) from ddld where trunc(SJ) =to_date(:RQ) and mc like '%自产%'");

            OracleCommand cmduery3 = new OracleCommand(sbSql3.ToString(), conn);
            cmduery3.Parameters.Add(":RQ", OracleType.DateTime).Value = riqi;
            OracleDataReader dr3 = cmduery3.ExecuteReader();
            while (dr3.Read())
            {
                ZCJT.DY80 = dr3.IsDBNull(0) ? null : (double?)dr3.GetDouble(0);
                ZCJT.ZJ86 = dr3.IsDBNull(1) ? null : (double?)dr3.GetDouble(1);

                ZCJT.ZJ64 = dr3.IsDBNull(2) ? null : (double?)dr3.GetDouble(2);
                ZCJT.ZJ42 = dr3.IsDBNull(3) ? null : (double?)dr3.GetDouble(3);
                ZCJT.XY25 = dr3.IsDBNull(4) ? null : (double?)dr3.GetDouble(4);
                ZCJT.PJLD = dr3.IsDBNull(5) ? null : (double?)dr3.GetDouble(5);
                ZCJT.LDJYX = dr3.IsDBNull(6) ? null : (double?)dr3.GetDouble(6);
            }
            result.Add(ZCJT);

            dr3.Close();



            StringBuilder sbSql4 = new StringBuilder();
            sbSql4.AppendLine("SELECT round(avg(C),3),round(avg(HUIFEN),3),round(avg(HUIFA),3),round(avg(S),3),round(avg(SHUIF),3),round(avg(M25),3),round(avg(M10),3),round(avg(QD),3) FROM DDJT WHERE trunc(SJ) =to_date(:RQ) and mc like '%外%'");
            OracleCommand cmduery4 = new OracleCommand(sbSql4.ToString(), conn);
            cmduery4.Parameters.Add(":RQ", OracleType.DateTime).Value = riqi;
            OracleDataReader dr4 = cmduery4.ExecuteReader();
            燃料指标内容项 WGJT = new 燃料指标内容项();
            while (dr4.Read())
            {

                WGJT.MC = "外购焦炭";
                WGJT.C = dr4.IsDBNull(0) ? null : (double?)dr4.GetDouble(0);
                WGJT.HUIFEN = dr4.IsDBNull(1) ? null : (double?)dr4.GetDouble(1);
                WGJT.HUIFA = dr4.IsDBNull(2) ? null : (double?)dr4.GetDouble(2);
                WGJT.S = dr4.IsDBNull(3) ? null : (double?)dr4.GetDouble(3);
                WGJT.SF = dr4.IsDBNull(4) ? null : (double?)dr4.GetDouble(4);
                WGJT.M40 = dr4.IsDBNull(5) ? null : (double?)dr4.GetDouble(5);
                WGJT.M10 = dr4.IsDBNull(6) ? null : (double?)dr4.GetDouble(6);
                WGJT.QD = dr4.IsDBNull(7) ? null : (double?)dr4.GetDouble(7);
            }
            dr4.Close();
            StringBuilder sbSql5 = new StringBuilder();
            sbSql5.AppendLine("select round(avg(DY80),3),round(avg(ZJ86),3),round(avg(ZJ64),3),round(avg(ZJ42),3),round(avg(XY25),3),round(avg(PJLD),3),round(avg(LDJYX),3) from ddld where trunc(SJ) =to_date(:RQ) and mc like '%外%'");

            OracleCommand cmduery5 = new OracleCommand(sbSql5.ToString(), conn);
            cmduery5.Parameters.Add(":RQ", OracleType.DateTime).Value = riqi;
            OracleDataReader dr5 = cmduery5.ExecuteReader();
            while (dr5.Read())
            {
                WGJT.DY80 = dr5.IsDBNull(0) ? null : (double?)dr5.GetDouble(0);
                WGJT.ZJ86 = dr5.IsDBNull(1) ? null : (double?)dr5.GetDouble(1);

                WGJT.ZJ64 = dr5.IsDBNull(2) ? null : (double?)dr5.GetDouble(2);
                WGJT.ZJ42 = dr5.IsDBNull(3) ? null : (double?)dr5.GetDouble(3);
                WGJT.XY25 = dr5.IsDBNull(4) ? null : (double?)dr5.GetDouble(4);
                WGJT.PJLD = dr5.IsDBNull(5) ? null : (double?)dr5.GetDouble(5);
                WGJT.LDJYX = dr5.IsDBNull(6) ? null : (double?)dr5.GetDouble(6);
            }
            result.Add(WGJT);
            dr5.Close();
            conn.Close();

        return result;
    }

        public 套筒窑内容 GetTTYReport(DateTime riqi)

        {
    
          
            套筒窑内容 result = new 套筒窑内容(riqi);
            StringBuilder sbSql = new StringBuilder();
            sbSql.AppendLine("select mc,round(avg(CaO),3),round(avg(SiO2),3),round(avg(Al2O3),3),round(avg(MgO),3),round(avg(S),3),round(avg(P),3),round(avg(HXD),3),round(avg(ZJ),3),round(avg(DY60),3),round(avg(ZJ64),3),round(avg(XY40),3) from ddtty where trunc(SJ) =to_date(:RQ) group by mc ");
            conn.Open();
            OracleCommand cmduery = new OracleCommand(sbSql.ToString(), conn);
            cmduery.Parameters.Add(":RQ", OracleType.DateTime).Value = riqi;
            OracleDataReader dr = cmduery.ExecuteReader();
            while (dr.Read())
            {
                套筒窑内容项 TTY = new 套筒窑内容项();
                TTY.CaO = dr.IsDBNull(1) ? null : (double?)dr.GetDouble(1);
                TTY.SiO2 = dr.IsDBNull(2) ? null : (double?)dr.GetDouble(2);

                TTY.Al2O3 = dr.IsDBNull(3) ? null : (double?)dr.GetDouble(3);
                TTY.MgO= dr.IsDBNull(4) ? null : (double?)dr.GetDouble(4);
                TTY.S = dr.IsDBNull(5) ? null : (double?)dr.GetDouble(5);
                TTY.P = dr.IsDBNull(6) ? null : (double?)dr.GetDouble(6);
                TTY.HXD = dr.IsDBNull(7) ? null : (double?)dr.GetDouble(7);
                TTY.ZJ = dr.IsDBNull(8) ? null : (double?)dr.GetDouble(8);
                TTY.DY60 = dr.IsDBNull(9) ? null : (double?)dr.GetDouble(9);
                TTY.ZJ64 = dr.IsDBNull(10) ? null : (double?)dr.GetDouble(10);
                TTY.XY40 = dr.IsDBNull(11) ? null : (double?)dr.GetDouble(11);
                result.Add(TTY);
            }
          
            dr.Close();
            conn.Close();

            return result; 
        
        
        
        
        
        
        
        
        
        
        
        
        
        }

    }
}
