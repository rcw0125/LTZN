using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;

namespace LTZN.查询
{
    class 风温考核
    {
        private double[,] defen = new double[6, 4];
        private double[,] banshu = new double[6, 4];
        private double[,] pingjunfen = new double[6, 4];
        public double jzfw1 = 1097;  //基准风温
        public double jzfw2 = 1085;
        public double jzfw3 = 1135;
        public double jzfw4 = 1020;
        public double jzfw5 = 1140;
        public double jzfw6 = 1140;

        private double[,] chanliang = new double[6, 4];
        private double[,] jiaobi = new double[6, 4];
        private double[,] meibi = new double[6, 4];
        private double[,] zonghejiaobi = new double[6, 4];
        private double[,] wandian= new double[6, 4];
        private double[,] dingwen = new double[6, 4];
        private double[,] fengwen = new double[6, 4];
        private double[,] si = new double[6, 4];
        private double[,] jt = new double[6, 4];
        private double[,] jd = new double[6, 4];
        private double[,] mei = new double[6, 4];
        public void display(HtmlDocument doc)
        {

            if (doc != null)
            {
                doc.OpenNew(true);
                doc.Write("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\" \"http://www.w3.org/TR/html4/loose.dtd\">");
                doc.Write("<HTML><head>");
                doc.Write("</head><body bgcolor=#EEEEFF topmargin=25 leftmargin=5>");
                doc.Write("<h1 align=center>分班指标</h1>");
                doc.Write("<hr width=80%>");
                doc.Write("<table  border=0 align=center><tr><td>");
                doc.Write("<table border=0 cellpadding=1 bgcolor=#cccccc  align=center>");
                doc.Write("<tr><td bgcolor=#C6FFC6 align=center height=30 width=56>高炉</td>");
                doc.Write("<td bgcolor=#C6FFC6 align=center height=30 width=56>班别</td>");
                doc.Write("<td bgcolor=#C6FFC6 align=center height=30 width=56>产量</td>");
                doc.Write("<td bgcolor=#C6FFC6 align=center height=30 width=56>焦比</td>");
                doc.Write("<td bgcolor=#C6FFC6 align=center height=30 width=56>煤比</td>");
                doc.Write("<td bgcolor=#C6FFC6 align=center height=30 width=56>综合焦比</td>");
                doc.Write("<td bgcolor=#C6FFC6 align=center height=30 width=56>晚点</td>");
                doc.Write("<td bgcolor=#C6FFC6 align=center height=30 width=56>顶温</td>");
                doc.Write("<td bgcolor=#C6FFC6 align=center height=30 width=56>风温</td>");
                doc.Write("<td bgcolor=#C6FFC6 align=center height=30 width=60>Si</td></tr>");
                doc.Write("</tr>");
                for (int i = 0; i < 6; i++)
                {
                    doc.Write("<tr>");
                    doc.Write("<td bgcolor=#D7D7FF height=24 align=center  rowspan=4>" + (i + 1) + "高炉</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >甲班</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(chanliang[i, 0]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(jiaobi[i, 0]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(meibi[i, 0]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(zonghejiaobi[i, 0]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml1(wandian[i, 0]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml1(dingwen[i, 0]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml1(fengwen[i, 0]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(si[i, 0]) + "</td>");
                    doc.Write("</tr>");
                    doc.Write("<tr>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >乙班</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(chanliang[i, 1]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(jiaobi[i, 1]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(meibi[i, 1]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(zonghejiaobi[i, 1]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml1(wandian[i, 1]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml1(dingwen[i, 1]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml1(fengwen[i, 1]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(si[i, 1]) + "</td>");
                    doc.Write("</tr>");
                    doc.Write("<tr>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >丙班</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(chanliang[i, 2]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(jiaobi[i, 2]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(meibi[i, 2]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(zonghejiaobi[i, 2]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml1(wandian[i, 2]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml1(dingwen[i, 2]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml1(fengwen[i, 2]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(si[i, 2]) + "</td>");
                    doc.Write("</tr>");
                    doc.Write("<tr>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >丁班</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(chanliang[i, 3]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(jiaobi[i, 3]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(meibi[i, 3]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(zonghejiaobi[i, 3]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml1(wandian[i, 3]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml1(dingwen[i, 3]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml1(fengwen[i, 3]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(si[i, 3]) + "</td>");
                    doc.Write("</tr>");
                }
                doc.Write("</table>");
                doc.Write("</td>");
                doc.Write("<td valign=top><span id=\"mingxi\">");

                doc.Write("</span></td>");
                doc.Write("</tr></table>");
                doc.Write("<hr width=80%>");
                doc.Write("</BODY></HTML>");
            }
        }
        private string convertToHtml(double d)
        {
            if (d > 0)
                return d.ToString("N2");
            else
                return "&nbsp;";

        }
        private string convertToHtml1(double d)
        {
            if (d > 0)
                return d.ToString("#####0");
            else
                return "&nbsp;";
        }
        public void getData(DateTime rq1, DateTime rq2)
        {
            OracleConnection cn = new OracleConnection();
            cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
            cn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = cn;
            cmd.Parameters.Add(":RQ1", OracleType.DateTime).Value = rq1;
            cmd.Parameters.Add(":RQ2", OracleType.DateTime).Value = rq2;
            cmd.CommandText = "select gaolu,GETBANCI1(ZDSJ),SUM(FELIANG),sum(wdsj) from ddluci where GETRQ1(zdsj)>=:RQ1 and GETRQ1(zdsj)<=:RQ2 and  zdsj is not null group by gaolu,GETBANCI1(ZDSJ)";
            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.IsDBNull(0) ? 0 : dr.GetInt32(0);
                    string banbie = dr.IsDBNull(1) ? "" : dr.GetString(1);
                    if (gaolu > 0 && gaolu < 7)
                    {
                        switch (banbie)
                        {
                            case "甲班":
                                chanliang[gaolu - 1, 0] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                                wandian[gaolu - 1, 0] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                                break;
                            case "乙班":
                                chanliang[gaolu - 1, 1] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                                wandian[gaolu - 1, 1] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                                break;
                            case "丙班":
                                chanliang[gaolu - 1, 2] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                                 wandian[gaolu - 1, 2] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                                break;
                            case "丁班":
                                chanliang[gaolu - 1, 3] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                                wandian[gaolu - 1, 3] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                                break;
                        }
                    }
                }
            }
            dr.Close();
            cmd.CommandText = "select gaolu,GETBANCI2(ZDSJ),round(avg(fesi),4) from ddluci where GETRQ2(zdsj)>=:RQ1 and GETRQ2(zdsj)<=:RQ2 and  zdsj is not null group by gaolu,GETBANCI2(ZDSJ)";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.IsDBNull(0) ? 0 : dr.GetInt32(0);
                    string banbie = dr.IsDBNull(1) ? "" : dr.GetString(1);
                    if (gaolu > 0 && gaolu <7)
                    {
                        switch (banbie)
                        {
                            case "甲班":
                                si[gaolu - 1, 0] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                                break;
                            case "乙班":
                                si[gaolu - 1, 1] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                                break;
                            case "丙班":
                                si[gaolu - 1, 2] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                                break;
                            case "丁班":
                                si[gaolu - 1, 3] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                                break;
                        }
                    }
                }
            }
            dr.Close();
            cmd.CommandText = "select gaolu,GETBANCI3(sj,hour),round(sum(pml),2),round(avg(ldwd),2),round(avg(rfwd),2) from rbcaozuo where GETRQ3(sj,hour)>=:RQ1 and GETRQ3(sj,hour)<=:RQ2 and ldwd>0 and rfwd>0 group by gaolu,GETBANCI3(sj,hour)";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.IsDBNull(0) ? 0 : dr.GetInt32(0);
                    string banbie = dr.IsDBNull(1) ? "" : dr.GetString(1);
                    if (gaolu > 0 && gaolu < 7)
                    {
                        switch (banbie)
                        {
                            case "甲班":
                                mei[gaolu - 1, 0] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2) *1000;
                                if (chanliang[gaolu - 1, 0] > 0)
                                    meibi[gaolu - 1, 0] = mei[gaolu - 1, 0] / chanliang[gaolu - 1, 0];
                                else
                                    meibi[gaolu - 1, 0] = 0;
                                dingwen[gaolu - 1, 0] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                                fengwen[gaolu - 1, 0] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                                break;
                            case "乙班":
                                mei[gaolu - 1, 1] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2) * 1000;
                                if (chanliang[gaolu - 1, 1] > 0)
                                    meibi[gaolu - 1, 1] = mei[gaolu - 1, 1] / chanliang[gaolu - 1, 1];
                                else
                                    meibi[gaolu - 1, 1] = 0;
                                dingwen[gaolu - 1, 1] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                                fengwen[gaolu - 1, 1] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                                break;
                            case "丙班":
                                mei[gaolu - 1, 2] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2) * 1000;
                                if (chanliang[gaolu - 1, 2] > 0)
                                    meibi[gaolu - 1, 2] = mei[gaolu - 1, 2] / chanliang[gaolu - 1, 2];
                                else
                                    meibi[gaolu - 1, 2] = 0;
                                dingwen[gaolu - 1, 2] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                                fengwen[gaolu - 1, 2] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                                break;
                            case "丁班":
                                mei[gaolu - 1, 3] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2) * 1000;
                                if (chanliang[gaolu - 1, 3] > 0)
                                    meibi[gaolu - 1, 3] = mei[gaolu - 1, 3] / chanliang[gaolu - 1, 3];
                                else
                                    meibi[gaolu - 1, 3] = 0;
                                dingwen[gaolu - 1, 3] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                                fengwen[gaolu - 1, 3] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                                break;
                        }
                    }
                }
            }
            dr.Close();
            cmd.CommandText = "select gaolu,GETBANCI4(T),sum(jt),sum(jd) from lt_liao where GETRQ4(T)>=:RQ1 and GETRQ4(T)<=:RQ2 group by gaolu,GETBANCI4(T)";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.IsDBNull(0) ? 0 : dr.GetInt32(0);
                    string banbie = dr.IsDBNull(1) ? "" : dr.GetString(1);
                    if (gaolu > 0 && gaolu < 7)
                    {
                        switch (banbie)
                        {
                            case "甲班":
                                jt[gaolu - 1, 0] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                                jd[gaolu - 1, 0] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                                if (chanliang[gaolu - 1, 0] > 0)
                                {
                                    jiaobi[gaolu - 1, 0] = (jt[gaolu - 1, 0] + jd[gaolu - 1, 0] * 0.9) / chanliang[gaolu - 1, 0];
                                    zonghejiaobi[gaolu - 1, 0] = (jt[gaolu - 1, 0] + jd[gaolu - 1, 0] * 0.9 + mei[gaolu - 1, 0] * 0.8) / chanliang[gaolu - 1, 0];
                                }
                                else
                                {
                                    jiaobi[gaolu - 1, 0] = 0;
                                    zonghejiaobi[gaolu - 1, 0] = 0;
                                }
                                    break;
                            case "乙班":
                                jt[gaolu - 1, 1] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                                jd[gaolu - 1, 1] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                                if (chanliang[gaolu - 1, 1] > 0)
                                {
                                    jiaobi[gaolu - 1, 1] = (jt[gaolu - 1, 1] + jd[gaolu - 1, 1] * 0.9) / chanliang[gaolu - 1, 0];
                                    zonghejiaobi[gaolu - 1, 1] = (jt[gaolu - 1, 1] + jd[gaolu - 1, 1] * 0.9 + mei[gaolu - 1, 1] * 0.8) / chanliang[gaolu - 1, 1];
                                }
                                else
                                {
                                    jiaobi[gaolu - 1, 1] = 0;
                                    zonghejiaobi[gaolu - 1, 1] = 0;
                                } 
                                break;
                            case "丙班":
                                jt[gaolu - 1, 2] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                                jd[gaolu - 1, 2] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                                if (chanliang[gaolu - 1, 2] > 0)
                                {
                                    jiaobi[gaolu - 1, 2] = (jt[gaolu - 1, 2] + jd[gaolu - 1, 2] * 0.9) / chanliang[gaolu - 1, 2];
                                    zonghejiaobi[gaolu - 1, 2] = (jt[gaolu - 1, 2] + jd[gaolu - 1, 2] * 0.9 + mei[gaolu - 1, 2] * 0.8) / chanliang[gaolu - 1, 2];

                                }
                                else
                                {
                                    jiaobi[gaolu - 1, 2] = 0;
                                    zonghejiaobi[gaolu - 1, 2] = 0;
                                }
                                break;
                            case "丁班":
                                jt[gaolu - 1, 3] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                                jd[gaolu - 1, 3] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                                if (chanliang[gaolu - 1, 3] > 0)
                                {
                                    jiaobi[gaolu - 1, 3] = (jt[gaolu - 1, 3] + jd[gaolu - 1, 3] * 0.9) / chanliang[gaolu - 1, 3];
                                    zonghejiaobi[gaolu - 1, 3] = (jt[gaolu - 1, 3] + jd[gaolu - 1, 3] * 0.9 + mei[gaolu - 1, 3] * 0.8) / chanliang[gaolu - 1, 3];
                                }
                                else
                                {
                                    jiaobi[gaolu - 1, 3] = 0;
                                    zonghejiaobi[gaolu - 1, 3] = 0;
                                }
                                break;
                        }
                    }
                }
            }
            dr.Close();
            cn.Close();
        
        }

        //public void getData(DateTime rq1, DateTime rq2)
        //{
        //    OracleConnection cn = new OracleConnection();
        //    cn.ConnectionString = Properties.Settings.Default.ConnectionString;
        //    cn.Open();
        //    OracleCommand cmd = new OracleCommand();
        //    cmd.Connection = cn;
        //    cmd.Parameters.Add(":RQ1", OracleType.DateTime).Value = rq1;
        //    cmd.Parameters.Add(":RQ2", OracleType.DateTime).Value = rq2;
        //    cmd.CommandText = "update rbcaozuo set banbie=(select 班别 from 高炉大班排班表 where 日期=trunc(rbcaozuo.sj) and 班次='白班') where to_number(to_char(sj,'HH24'))>=9 and to_number(to_char(sj,'HH24'))<=19 and gaolu>2  and trunc(sj+4/24)>=:RQ1 and trunc(sj+4/24)<=:RQ2";
        //    cmd.ExecuteNonQuery();
        //    cmd.CommandText = "update rbcaozuo set banbie=(select 班别 from 高炉大班排班表 where 日期=trunc(rbcaozuo.sj) and 班次='夜班') where to_number(to_char(sj,'HH24'))<9  and gaolu <>1  and trunc(sj+4/24)>=:RQ1 and trunc(sj+4/24)<=:RQ2";
        //    cmd.ExecuteNonQuery();
        //    cmd.CommandText = "update rbcaozuo set banbie=(select 班别 from 高炉大班排班表 where 日期=trunc(rbcaozuo.sj+1) and 班次='夜班') where to_number(to_char(sj,'HH24'))>19 and gaolu>2  and trunc(sj+4/24)>=:RQ1 and trunc(sj+4/24)<=:RQ2";
        //    cmd.ExecuteNonQuery();
        //    cmd.CommandText = "update rbcaozuo set banbie=(select 班别 from 高炉大班排班表 where 日期=trunc(rbcaozuo.sj) and 班次='白班') where to_number(to_char(sj,'HH24'))>=9 and to_number(to_char(sj,'HH24'))<=20 and gaolu=2  and trunc(sj+4/24)>=:RQ1 and trunc(sj+4/24)<=:RQ2";
        //    cmd.ExecuteNonQuery();
        //    cmd.CommandText = "update rbcaozuo set banbie=(select 班别 from 高炉大班排班表 where 日期=trunc(rbcaozuo.sj+1) and 班次='夜班') where to_number(to_char(sj,'HH24'))>20 and gaolu=2  and trunc(sj+4/24)>=:RQ1 and trunc(sj+4/24)<=:RQ2";
        //    cmd.ExecuteNonQuery();
        //    cmd.CommandText = "update rbcaozuo set banbie=(select 班别 from 高炉排班表 where 日期=trunc(rbcaozuo.sj) and 班次='夜班') where to_number(to_char(sj,'HH24'))>=1 and to_number(to_char(sj,'HH24'))<=8 and gaolu=1  and trunc(sj+4/24)>=:RQ1 and trunc(sj+4/24)<=:RQ2";
        //    cmd.ExecuteNonQuery();
        //    cmd.CommandText = "update rbcaozuo set banbie=(select 班别 from 高炉排班表 where 日期=trunc(rbcaozuo.sj) and 班次='白班') where to_number(to_char(sj,'HH24'))>=9 and to_number(to_char(sj,'HH24'))<=16 and gaolu=1  and trunc(sj+4/24)>=:RQ1 and trunc(sj+4/24)<=:RQ2";
        //    cmd.ExecuteNonQuery();
        //    cmd.CommandText = "update rbcaozuo set banbie=(select 班别 from 高炉排班表 where 日期=trunc(rbcaozuo.sj) and 班次='中班') where to_number(to_char(sj,'HH24'))>=17 and gaolu=1  and trunc(sj+4/24)>=:RQ1 and trunc(sj+4/24)<=:RQ2";
        //    cmd.ExecuteNonQuery();
        //    cmd.CommandText = "update rbcaozuo set banbie=(select 班别 from 高炉排班表 where 日期=trunc(rbcaozuo.sj-1) and 班次='中班') where to_number(to_char(sj,'HH24'))=0 and gaolu=1  and trunc(sj+4/24)>=:RQ1 and trunc(sj+4/24)<=:RQ2";
        //    cmd.ExecuteNonQuery();      
        //    cmd.Parameters.Add(":JZFW", OracleType.Number).Value = jzfw1;
        //    cmd.CommandText = "select banbie,sum((rfwd-:JZFW)*0.2),count(*) from(select trunc(sj-1/24),banbie,trunc(avg(rfwd),2) as rfwd from rbcaozuo where rfwd>0 and gaolu=1 and trunc(sj-1/24)>=:RQ1 and trunc(sj-1/24)<=:RQ2 and not EXISTS(select * from 休风 where sj>=时间 and ((间隔>=240 and sj<=时间+间隔*1.5/1440) or (间隔<240 and sj<=时间+(间隔+120)/1440)) and 高炉=gaolu)  group by trunc(sj-1/24),banbie) group by banbie";
        //    OracleDataReader dr = cmd.ExecuteReader();
        //    while (dr.Read())
        //    {
        //        if (!dr.IsDBNull(0))
        //        {
        //            string banbie = dr.IsDBNull(0) ? "" : dr.GetString(0);
        //            switch (banbie)
        //            {
        //                case "甲班":
        //                    defen[0, 0] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
        //                    banshu[0, 0] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
        //                    break;
        //                case "乙班":
        //                    defen[0, 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
        //                    banshu[0, 1] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
        //                    break;
        //                case "丙班":
        //                    defen[0, 2] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
        //                    banshu[0, 2] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
        //                    break;
        //                case "丁班":
        //                    defen[0, 3] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
        //                    banshu[0, 3] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
        //                    break;
        //            }
        //        }
        //    }
        //    dr.Close();
        //    cmd.Parameters[":JZFW"].Value = jzfw2;
        //    cmd.CommandText = "select banbie,sum((rfwd-:JZFW)*0.2),count(*) from(select trunc(sj+3/24),banbie,trunc(avg(rfwd),2) as rfwd from rbcaozuo where rfwd>0 and gaolu=2 and trunc(sj+3/24)>=:RQ1 and trunc(sj+3/24)<=:RQ2 and not EXISTS(select * from 休风 where sj>=时间 and ((间隔>=240 and sj<=时间+间隔*1.5/1440) or (间隔<240 and sj<=时间+(间隔+120)/1440)) and 高炉=gaolu) group by trunc(sj+3/24),banbie) group by banbie";
        //    dr = cmd.ExecuteReader();
        //    while (dr.Read())
        //    {
        //        if (!dr.IsDBNull(0))
        //        {
        //            string banbie = dr.IsDBNull(0) ? "" : dr.GetString(0);
        //            switch (banbie)
        //            {
        //                case "甲班":
        //                    defen[1, 0] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
        //                    banshu[1, 0] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
        //                    break;
        //                case "乙班":
        //                    defen[1, 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
        //                    banshu[1, 1] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
        //                    break;
        //                case "丙班":
        //                    defen[1, 2] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
        //                    banshu[1, 2] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
        //                    break;
        //                case "丁班":
        //                    defen[1, 3] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
        //                    banshu[1, 3] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
        //                    break;
        //            }
        //        }
        //    }
        //    dr.Close();
        //    cmd.Parameters[":JZFW"].Value = jzfw3;
        //    cmd.CommandText = "select banbie,sum((rfwd-:JZFW)*0.2),count(*) from(select trunc(sj+4/24),banbie,trunc(avg(rfwd),2) as rfwd from rbcaozuo where rfwd>0 and gaolu=3 and trunc(sj+4/24)>=:RQ1 and trunc(sj+4/24)<=:RQ2 and not EXISTS(select * from 休风 where sj>=时间 and ((间隔>=240 and sj<=时间+间隔*1.5/1440) or (间隔<240 and sj<=时间+(间隔+120)/1440)) and 高炉=gaolu)  group by trunc(sj+4/24),banbie) group by banbie";
        //    dr = cmd.ExecuteReader();
        //    while (dr.Read())
        //    {
        //        if (!dr.IsDBNull(0))
        //        {

        //            string banbie = dr.IsDBNull(0) ? "" : dr.GetString(0);
        //            switch (banbie)
        //            {
        //                case "甲班":
        //                    defen[2, 0] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
        //                    banshu[2, 0] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
        //                    break;
        //                case "乙班":
        //                    defen[2, 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
        //                    banshu[2, 1] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
        //                    break;
        //                case "丙班":
        //                    defen[2, 2] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
        //                    banshu[2, 2] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
        //                    break;
        //                case "丁班":
        //                    defen[2, 3] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
        //                    banshu[2, 3] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
        //                    break;
        //            }
        //        }
        //    }
        //    dr.Close();
        //    cmd.Parameters[":JZFW"].Value = jzfw4;
        //    cmd.CommandText = "select banbie,sum((rfwd-:JZFW)*0.2),count(*) from(select trunc(sj+4/24),banbie,trunc(avg(rfwd),2) as rfwd from rbcaozuo where rfwd>0 and gaolu=4 and trunc(sj+4/24)>=:RQ1 and trunc(sj+4/24)<=:RQ2 and not EXISTS(select * from 休风 where sj>=时间 and ((间隔>=240 and sj<=时间+间隔*1.5/1440) or (间隔<240 and sj<=时间+(间隔+120)/1440)) and 高炉=gaolu)  group by trunc(sj+4/24),banbie) group by banbie";
        //    dr = cmd.ExecuteReader();
        //    while (dr.Read())
        //    {
        //        if (!dr.IsDBNull(0))
        //        {
        //            string banbie = dr.IsDBNull(0) ? "" : dr.GetString(0);
        //            switch (banbie)
        //            {
        //                case "甲班":
        //                    defen[3, 0] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
        //                    banshu[3, 0] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
        //                    break;
        //                case "乙班":
        //                    defen[3, 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
        //                    banshu[3, 1] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
        //                    break;
        //                case "丙班":
        //                    defen[3, 2] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
        //                    banshu[3, 2] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
        //                    break;
        //                case "丁班":
        //                    defen[3, 3] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
        //                    banshu[3, 3] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
        //                    break;
        //            }
        //        }
        //    }
        //    dr.Close();
        //    cmd.Parameters[":JZFW"].Value = jzfw5;
        //    cmd.CommandText = "select banbie,sum((rfwd-:JZFW)*0.2),count(*) from(select trunc(sj+4/24),banbie,trunc(avg(rfwd),2) as rfwd from rbcaozuo where rfwd>0 and gaolu=5 and trunc(sj+4/24)>=:RQ1 and trunc(sj+4/24)<=:RQ2 and not EXISTS(select * from 休风 where sj>=时间 and ((间隔>=240 and sj<=时间+间隔*1.5/1440) or (间隔<240 and sj<=时间+(间隔+120)/1440)) and 高炉=gaolu)  group by trunc(sj+4/24),banbie) group by banbie";
        //    dr = cmd.ExecuteReader();
        //    while (dr.Read())
        //    {
        //        if (!dr.IsDBNull(0))
        //        {
        //            string banbie = dr.IsDBNull(0) ? "" : dr.GetString(0);
        //            switch (banbie)
        //            {
        //                case "甲班":
        //                    defen[4, 0] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
        //                    banshu[4, 0] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
        //                    break;
        //                case "乙班":
        //                    defen[4, 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
        //                    banshu[4, 1] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
        //                    break;
        //                case "丙班":
        //                    defen[4, 2] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
        //                    banshu[4, 2] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
        //                    break;
        //                case "丁班":
        //                    defen[4, 3] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
        //                    banshu[4, 3] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
        //                    break;
        //            }
        //        }
        //    }
        //    dr.Close();
        //    cn.Close();
        //    for (int i = 0; i < 5; i++)
        //        for (int j = 0; j < 4; j++)
        //            if (banshu[i, j] > 0)
        //                pingjunfen[i, j] = defen[i, j] / banshu[i, j];
        //}

    }
}
