using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;
using LTZN.Repository;
using LTZN.CalFramework;

namespace LTZN.查询
{
    public class 上报厂调:QueryModel
    {
        public 上报厂调()
            : base("上报厂调")
        {

        }

        public void getData(DateTime riqi)
        {
            this.myModel.Init();

            conn.Open();
            String sqlstr = "select 高炉,机烧矿,竖球,本溪矿,生料,工艺称,煤粉,自产湿焦,落地湿焦 from 原料消耗 where 日期=:rq";
            OracleCommand cmds1 = new OracleCommand(sqlstr, conn);
            cmds1.Parameters.Add(":rq", OracleType.DateTime).Value = riqi;

            OracleDataReader dr = cmds1.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("高炉{0}.机烧", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                    this.SetValue(string.Format("高炉{0}.球团", gaolu), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                    this.SetValue(string.Format("高炉{0}.本溪", gaolu), dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                    this.SetValue(string.Format("高炉{0}.生料", gaolu), dr.IsDBNull(4) ? 0 : dr.GetDouble(4));
                    this.SetValue(string.Format("高炉{0}.煤粉", gaolu), dr.IsDBNull(6) ? 0 : dr.GetDouble(6));
                    this.SetValue(string.Format("高炉{0}.自产湿焦", gaolu), dr.IsDBNull(7) ? 0 : dr.GetDouble(7));
                    this.SetValue(string.Format("高炉{0}.落地湿焦", gaolu), dr.IsDBNull(8) ? 0 : dr.GetDouble(8));
                }

            }
            dr.Close();

            cmds1.CommandText = "select 高炉,SUM(机烧矿),SUM(竖球),SUM(本溪矿),SUM(生料),SUM(工艺称),SUM(煤粉),SUM(自产湿焦),SUM(落地湿焦) from 原料消耗 where trunc(日期,'MONTH')=trunc(:rq,'MONTH') and 日期<=:rq group by 高炉";
            dr = cmds1.ExecuteReader();
            while (dr.Read())
            {

                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("高炉{0}.机烧累计", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                    this.SetValue(string.Format("高炉{0}.球团累计", gaolu), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                    this.SetValue(string.Format("高炉{0}.本溪累计", gaolu), dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                    this.SetValue(string.Format("高炉{0}.生料累计", gaolu), dr.IsDBNull(4) ? 0 : dr.GetDouble(4));
                    this.SetValue(string.Format("高炉{0}.煤粉累计", gaolu), dr.IsDBNull(6) ? 0 : dr.GetDouble(6));
                    this.SetValue(string.Format("高炉{0}.自产湿焦累计", gaolu), dr.IsDBNull(7) ? 0 : dr.GetDouble(7));
                    this.SetValue(string.Format("高炉{0}.落地湿焦累计", gaolu), dr.IsDBNull(8) ? 0 : dr.GetDouble(8));
                }

            }
            dr.Close();

            cmds1.CommandText = "select gaolu,sum(feliang) from ddluci where trunc(zdsj)=:rq and dksj is not null group by gaolu";
            dr = cmds1.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("高炉{0}.全铁产量", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();

            cmds1.CommandText = "select gaolu,sum(feliang) from ddluci where trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq  and dksj is not null group by gaolu";
            dr = cmds1.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("高炉{0}.全铁产量累计", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();


            cmds1.CommandText = "select gaolu,sum(FELIANG) from ddluci where quchu='铸造' and trunc(zdsj)=:rq  and dksj is not null group by gaolu";
            dr = cmds1.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("高炉{0}.铸造铁", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();

            cmds1.CommandText = "select gaolu,sum(FELIANG) from ddluci where quchu='铸造' and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq  and dksj is not null group by gaolu";
            dr = cmds1.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("高炉{0}.铸造铁累计", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();


            cmds1.CommandText = "select gaolu,sum(FELIANG) from ddluci where quchu='炼钢' and trunc(zdsj)=:rq  and dksj is not null group by gaolu";
            dr = cmds1.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("高炉{0}.送炼钢铁", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();

            cmds1.CommandText = "select gaolu,sum(FELIANG) from ddluci where quchu='炼钢' and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq  and dksj is not null group by gaolu";
            dr = cmds1.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("高炉{0}.送炼钢铁累计", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();

            cmds1.CommandText = "select gaolu,sum(FELIANG) from ddluci where FESi<=0.7 and trunc(zdsj)=:rq  and dksj is not null group by gaolu";
            dr = cmds1.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("高炉{0}.Si小于0点7", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();

            cmds1.CommandText = "select gaolu,sum(FELIANG) from ddluci where FESi<=0.7 and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq  and dksj is not null group by gaolu";
            dr = cmds1.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("高炉{0}.Si小于0点7累计", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();

            cmds1.CommandText = "select trunc(avg(shuif),2) from ddjt where trunc(sj)=:RQ and mc in ('自产焦夜','自产焦白','自产焦中')";
            dr = cmds1.ExecuteReader();
            if (dr.Read())
                this.SetValue("自产湿焦水份", dr.IsDBNull(0) ? 0 : dr.GetDouble(0));
            dr.Close();

            cmds1.CommandText = "select trunc(avg(shuif),2) from ddjt where trunc(sj)=:RQ and mc='外进焦'";
            dr = cmds1.ExecuteReader();
            if (dr.Read())
                this.SetValue("落地湿焦水份", dr.IsDBNull(0) ? 0 : dr.GetDouble(0));
            dr.Close();


            cmds1.CommandText = "select trunc(avg(shuif),2) from ddjt where trunc(sj,'MONTH')=trunc(:rq,'MONTH') and trunc(sj)<=:RQ and mc in ('自产焦夜','自产焦白','自产焦中')";
            dr = cmds1.ExecuteReader();
            if (dr.Read())
                this.SetValue("自产湿焦水份累计", dr.IsDBNull(0) ? 0 : dr.GetDouble(0));
            dr.Close();
            cmds1.CommandText = "select trunc(avg(shuif),2) from ddjt where  trunc(sj,'MONTH')=trunc(:rq,'MONTH') and trunc(sj)<=:RQ and mc='外进焦'";
            dr = cmds1.ExecuteReader();
            if (dr.Read())
                this.SetValue("落地湿焦水份累计", dr.IsDBNull(0) ? 0 : dr.GetDouble(0));
            dr.Close();
            conn.Close();
        }

    }

    //class 上报厂调
    //{
    //    private double[] 全铁产量 =new double[6];
    //    private double[] SI小于07 = new double[6];
    //    private double[] 送炼钢铁 = new double[6];
    //    private double[] 铸造铁 = new double[6];
    //    private double[] 机烧 = new double[6];
    //    private double[] 球团 = new double[6];
    //    private double[] 本溪 = new double[6];
    //    private double[] 生料 = new double[6];
    //    private double[] 工艺称焦炭 = new double[6];
    //    private double[] 煤粉 = new double[6];
    //    private double[] 焦比 = new double[6];
    //    private double[] 煤比 = new double[6];
    //    private double[] 自产湿焦 = new double[6];
    //    private double[] 落地湿焦 = new double[6];
    //    private double[] 入炉焦炭 = new double[6];

    //    private double[] 累计全铁产量 = new double[6];
    //    private double[] 累计SI小于07 = new double[6];
    //    private double[] 累计送炼钢铁 = new double[6];
    //    private double[] 累计铸造铁 = new double[6];
    //    private double[] 累计机烧 = new double[6];
    //    private double[] 累计球团 = new double[6];
    //    private double[] 累计本溪 = new double[6];
    //    private double[] 累计生料 = new double[6];
    //    private double[] 累计工艺称焦炭 = new double[6];
    //    private double[] 累计煤粉 = new double[6];
    //    private double[] 累计焦比 = new double[6];
    //    private double[] 累计煤比 = new double[6];
    //    private double[] 累计自产湿焦 = new double[6];
    //    private double[] 累计落地湿焦 = new double[6];
    //    private double[] 累计入炉焦炭 = new double[6];

    //    private string convertToHtml(double d)
    //    {
    //        if (d > 0)
    //            return d.ToString("#######0.00");
    //        else
    //            return "&nbsp;";
    //    }
    //    private string convertToHtml(double d, string format)
    //    {
    //        if (d > 0)
    //            return d.ToString(format);
    //        else
    //            return "&nbsp;";
    //    }

    //    public void display(HtmlDocument doc)
    //    {
    //        if (doc != null)
    //        {
    //            doc.OpenNew(true);
    //            doc.Write("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\" \"http://www.w3.org/TR/html4/loose.dtd\">");
    //            doc.Write("<HTML><body bgcolor=#EEEEFF topmargin=25 leftmargin=5>");
    //            doc.Write("     <h1 align=center>厂调指标查询</h1>");
    //            doc.Write("        <hr width=80%>");
    //            doc.Write("        <table border=0 cellpadding=1 bgcolor=#cccccc height=151 align=center>");
    //            doc.Write("          <tr>");
    //            doc.Write("            <td bgcolor=#C6FFC6 align=center height=28 width=40 align=center>单位</td>");
    //            doc.Write("            <td bgcolor=#C6FFC6 align=center height=28 width=40 align=center>项目</td>");
    //            doc.Write("            <td bgcolor=#C6FFC6 align=center height=28 width=66>全铁产量</td>");
    //            doc.Write("            <td bgcolor=#C6FFC6 align=center height=28 width=64>Si≤0.7</td>");
    //            doc.Write("            <td bgcolor=#C6FFC6 align=center height=28 width=64>送炼钢铁</td>");
    //            doc.Write("            <td bgcolor=#C6FFC6 align=center height=28 width=56>铸造铁</td>");
    //            doc.Write("            <td bgcolor=#C6FFC6 align=center height=28 width=56>机烧</td>");
    //            doc.Write("            <td bgcolor=#C6FFC6 align=center height=28 width=56>球团</td>");
    //            doc.Write("            <td bgcolor=#C6FFC6 align=center height=28 width=56>本溪</td>");
    //            doc.Write("            <td bgcolor=#C6FFC6 align=center height=28 width=56>生料</td>");
    //            doc.Write("            <td bgcolor=#C6FFC6 align=center height=28 width=56>焦比</td>");
    //            doc.Write("            <td bgcolor=#C6FFC6 align=center height=28 width=56>煤比</td>");
    //            doc.Write("          </tr>");

    //            for (int i = 0; i < 5; i++)
    //            {

    //                doc.Write("          <tr>");
    //                doc.Write("            <td height=24 bgcolor=#D7D7FF align=center>" + (i + 1) + "#</td>");
    //                doc.Write("            <td height=24 bgcolor=#D7D7FF align=center >本日</td>");
    //                doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right >" + convertToHtml(全铁产量[i]) + "</td>");
    //                doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right >" + convertToHtml(SI小于07[i]) + "</td>");
    //                doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right >" + convertToHtml(送炼钢铁[i]) + "</td>");
    //                doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right >" + convertToHtml(铸造铁[i]) + "</td>");
    //                doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right >" + convertToHtml(机烧[i]) + "</td>");
    //                doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right >" + convertToHtml(球团[i]) + "</td>");
    //                doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right >" + convertToHtml(本溪[i]) + "</td>");
    //                doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right >" + convertToHtml(生料[i]) + "</td>");
    //                doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right >" + convertToHtml(焦比[i], "###0") + "</td>");
    //                doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right >" + convertToHtml(煤比[i], "###0") + "</td>");
    //                doc.Write("          </tr><tr>");
    //                doc.Write("            <td height=24 bgcolor=#D7D7FF align=center>" + (i + 1) + "#</td>");
    //                doc.Write("            <td height=24 bgcolor=#D7D7FF align=center>累计</td>");
    //                doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right>" + convertToHtml(累计全铁产量[i]) + "</td>");
    //                doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right>" + convertToHtml(累计SI小于07[i]) + "</td>");
    //                doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right>" + convertToHtml(累计送炼钢铁[i]) + "</td>");
    //                doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right>" + convertToHtml(累计铸造铁[i]) + "</td>");
    //                doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right>" + convertToHtml(累计机烧[i]) + "</td>");
    //                doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right>" + convertToHtml(累计球团[i]) + "</td>");
    //                doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right>" + convertToHtml(累计本溪[i]) + "</td>");
    //                doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right>" + convertToHtml(累计生料[i]) + "</td>");
    //                doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right>" + convertToHtml(累计焦比[i], "###0") + "</td>");
    //                doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right>" + convertToHtml(累计煤比[i], "###0") + "</td>");
    //                doc.Write("           </tr>");
    //            }

    //            doc.Write("          <tr>");
    //            doc.Write("            <td height=24 bgcolor=#D7D7FF align=center>全厂</td>");
    //            doc.Write("            <td height=24 bgcolor=#D7D7FF align=center>本日</td>");
    //            doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right >" + convertToHtml(全铁产量[5]) + "</td>");
    //            doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right >" + convertToHtml(SI小于07[5]) + "</td>");
    //            doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right >" + convertToHtml(送炼钢铁[5]) + "</td>");
    //            doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right >" + convertToHtml(铸造铁[5]) + "</td>");
    //            doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right >" + convertToHtml(机烧[5]) + "</td>");
    //            doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right >" + convertToHtml(球团[5]) + "</td>");
    //            doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right >" + convertToHtml(本溪[5]) + "</td>");
    //            doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right >" + convertToHtml(生料[5]) + "</td>");
    //            doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right >" + convertToHtml(焦比[5], "###0") + "</td>");
    //            doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right >" + convertToHtml(煤比[5], "###0") + "</td>");
    //            doc.Write("          </tr>");
    //            doc.Write("          <tr>");
    //            doc.Write("            <td height=24 bgcolor=#D7D7FF align=center>全厂</td>");
    //            doc.Write("            <td height=24 bgcolor=#D7D7FF align=center>累计</td>");
    //            doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right>" + convertToHtml(累计全铁产量[5]) + "</td>");
    //            doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right>" + convertToHtml(累计SI小于07[5]) + "</td>");
    //            doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right>" + convertToHtml(累计送炼钢铁[5]) + "</td>");
    //            doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right>" + convertToHtml(累计铸造铁[5]) + "</td>");
    //            doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right>" + convertToHtml(累计机烧[5]) + "</td>");
    //            doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right>" + convertToHtml(累计球团[5]) + "</td>");
    //            doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right>" + convertToHtml(累计本溪[5]) + "</td>");
    //            doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right>" + convertToHtml(累计生料[5]) + "</td>");
    //            doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right>" + convertToHtml(累计焦比[5], "###0") + "</td>");
    //            doc.Write("            <td height=24 bgcolor=#FFFFD7 align=right>" + convertToHtml(累计煤比[5], "###0") + "</td>");
    //            doc.Write("          </tr>");
    //            doc.Write("        </table>");
    //            doc.Write("</BODY></HTML>");
    //        }
    //        Program.Log(doc.Body.InnerHtml);
    //    }

    //    public void getData(DateTime riqi)
    //    {
    //        OracleConnection cn = new OracleConnection();
    //        cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
    //        cn.Open();
    //        String sqlstr = "select 高炉,机烧矿,竖球,本溪矿,生料,工艺称,煤粉,自产湿焦,落地湿焦 from 原料消耗 where 日期=:rq";
    //        OracleCommand cmds1 = new OracleCommand(sqlstr, cn);
    //        cmds1.Parameters.Add(":rq", OracleType.DateTime).Value = riqi;

    //        OracleDataReader dr = cmds1.ExecuteReader();
    //        while (dr.Read())
    //        {
    //            if (!dr.IsDBNull(0))
    //            {
    //                int gaolu = dr.GetInt32(0);
    //                机烧[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
    //                球团[gaolu - 1] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
    //                本溪[gaolu - 1] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
    //                生料[gaolu - 1] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
    //               // 工艺称焦炭[gaolu - 1] = dr.IsDBNull(5) ? 0 : dr.GetDouble(5);
    //                煤粉[gaolu - 1] = dr.IsDBNull(6) ? 0 : dr.GetDouble(6);
    //                自产湿焦[gaolu - 1] = dr.IsDBNull(7) ? 0 : dr.GetDouble(7);
    //                落地湿焦[gaolu - 1] = dr.IsDBNull(8) ? 0 : dr.GetDouble(8);
    //            }

    //        }
    //        dr.Close();

    //        cmds1.CommandText = "select 高炉,SUM(机烧矿),SUM(竖球),SUM(本溪矿),SUM(生料),SUM(工艺称),SUM(煤粉),SUM(自产湿焦),SUM(落地湿焦) from 原料消耗 where trunc(日期,'MONTH')=trunc(:rq,'MONTH') and 日期<=:rq group by 高炉";
    //        dr = cmds1.ExecuteReader();
    //        while (dr.Read())
    //        {

    //            if (!dr.IsDBNull(0))
    //            {
    //                int gaolu = dr.GetInt32(0);
    //                累计机烧[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
    //                累计球团[gaolu - 1] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
    //                累计本溪[gaolu - 1] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
    //                累计生料[gaolu - 1] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
    //               // 累计工艺称焦炭[gaolu - 1] = dr.IsDBNull(5) ? 0 : dr.GetDouble(5);
    //                累计煤粉[gaolu - 1] = dr.IsDBNull(6) ? 0 : dr.GetDouble(6);
    //                累计自产湿焦[gaolu - 1] = dr.IsDBNull(7) ? 0 : dr.GetDouble(7);
    //                累计落地湿焦[gaolu - 1] = dr.IsDBNull(8) ? 0 : dr.GetDouble(8);
    //            }

    //        }
    //        dr.Close();

    //        cmds1.CommandText = "select gaolu,sum(feliang) from ddluci where trunc(zdsj)=:rq and dksj is not null group by gaolu";      
    //        dr = cmds1.ExecuteReader();
    //        while (dr.Read())
    //        {
    //            if (!dr.IsDBNull(0))
    //            {
    //                int gaolu = dr.GetInt32(0);
    //                全铁产量[gaolu - 1] = (dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
    //            }
    //        }
    //        dr.Close();

    //        cmds1.CommandText = "select gaolu,sum(feliang) from ddluci where trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq  and dksj is not null group by gaolu";
    //        dr = cmds1.ExecuteReader();
    //        while (dr.Read())
    //        {
    //            if (!dr.IsDBNull(0))
    //            {
    //                int gaolu = dr.GetInt32(0);
    //                累计全铁产量[gaolu - 1] = (dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
    //            }
    //        }
    //        dr.Close();


    //        cmds1.CommandText= "select gaolu,sum(FELIANG) from ddluci where quchu='铸造' and trunc(zdsj)=:rq  and dksj is not null group by gaolu";
    //        dr = cmds1.ExecuteReader();
    //        while (dr.Read())
    //        {
    //            if (!dr.IsDBNull(0))
    //            {
    //                int gaolu = dr.GetInt32(0);
    //                铸造铁[gaolu - 1] = (dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
    //            }
    //        }
    //        dr.Close();

    //        cmds1.CommandText = "select gaolu,sum(FELIANG) from ddluci where quchu='铸造' and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq  and dksj is not null group by gaolu";
    //        dr = cmds1.ExecuteReader();
    //        while (dr.Read())
    //        {
    //            if (!dr.IsDBNull(0))
    //            {
    //                int gaolu = dr.GetInt32(0);
    //                累计铸造铁[gaolu - 1] = (dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
    //            }
    //        }
    //        dr.Close();


    //        cmds1.CommandText = "select gaolu,sum(FELIANG) from ddluci where quchu='炼钢' and trunc(zdsj)=:rq  and dksj is not null group by gaolu";
    //        dr = cmds1.ExecuteReader();
    //        while (dr.Read())
    //        {
    //            if (!dr.IsDBNull(0))
    //            {
    //                int gaolu = dr.GetInt32(0);
    //                送炼钢铁[gaolu - 1] = (dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
    //            }
    //        }
    //        dr.Close();

    //        cmds1.CommandText = "select gaolu,sum(FELIANG) from ddluci where quchu='炼钢' and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq  and dksj is not null group by gaolu";
    //        dr = cmds1.ExecuteReader();
    //        while (dr.Read())
    //        {
    //            if (!dr.IsDBNull(0))
    //            {
    //                int gaolu = dr.GetInt32(0);
    //                累计送炼钢铁[gaolu - 1] = (dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
    //            }
    //        }
    //        dr.Close();

    //        cmds1.CommandText = "select gaolu,sum(FELIANG) from ddluci where FESi<=0.7 and trunc(zdsj)=:rq  and dksj is not null group by gaolu";
    //        dr = cmds1.ExecuteReader();
    //        while (dr.Read())
    //        {
    //            if (!dr.IsDBNull(0))
    //            {
    //                int gaolu = dr.GetInt32(0);
    //                SI小于07[gaolu - 1] = (dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
    //            }
    //        }
    //        dr.Close();

    //        cmds1.CommandText = "select gaolu,sum(FELIANG) from ddluci where FESi<=0.7 and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq  and dksj is not null group by gaolu";
    //        dr = cmds1.ExecuteReader();
    //        while (dr.Read())
    //        {
    //            if (!dr.IsDBNull(0))
    //            {
    //                int gaolu = dr.GetInt32(0);
    //                累计SI小于07[gaolu - 1] = (dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
    //            }
    //        }
    //        dr.Close();

    //        double 自产焦水分 = 0;
    //        double 落地焦水分 = 0;
    //        //double 自产焦 = 0;
    //        //double 落地焦 = 0;

    //        double 累计自产焦平均水分 = 0;
    //        double 累计落地焦平均水分 = 0;
    //        //double 累计自产焦 = 0;
    //        //double 累计落地焦 = 0;

    //        //double 入炉焦炭 = 0;
    //        //double 累计入炉焦炭 = 0;

    //        cmds1.CommandText = "select trunc(avg(shuif),2) from ddjt where trunc(sj)=:RQ and mc in ('自产焦夜','自产焦白','自产焦中')";
    //        dr = cmds1.ExecuteReader();
    //        if (dr.Read())
    //            自产焦水分 = (dr.IsDBNull(0) ? 0 : dr.GetDouble(0));
    //        dr.Close();
    //        cmds1.CommandText = "select trunc(avg(shuif),2) from ddjt where trunc(sj)=:RQ and mc='外进焦'";
    //        dr = cmds1.ExecuteReader();
    //        if (dr.Read())
    //            落地焦水分 = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
    //        dr.Close();
    //        //cmds1.CommandText = "select 自产湿焦,落地湿焦 from 全厂日消耗 where P01日期=:rq";
    //        //dr = cmds1.ExecuteReader();
    //        //if (dr.Read())
    //        //{
    //        //    自产焦 = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
    //        //    落地焦 = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
    //        //}
    //        //dr.Close();

    //        cmds1.CommandText = "select trunc(avg(shuif),2) from ddjt where trunc(sj,'MONTH')=trunc(:rq,'MONTH') and trunc(sj)<=:RQ and mc in ('自产焦夜','自产焦白','自产焦中')";
    //        dr = cmds1.ExecuteReader();
    //        if (dr.Read())
    //            累计自产焦平均水分= dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
    //        dr.Close();
    //        cmds1.CommandText = "select trunc(avg(shuif),2) from ddjt where  trunc(sj,'MONTH')=trunc(:rq,'MONTH') and trunc(sj)<=:RQ and mc='外进焦'";
    //        dr = cmds1.ExecuteReader();
    //        if (dr.Read())
    //            累计落地焦平均水分 = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
    //        dr.Close();
    //        //cmds1.CommandText = "select SUM(自产湿焦),sum(落地湿焦) from 全厂日消耗 where  trunc(P01日期,'MONTH')=trunc(:rq,'MONTH') and P01日期<=:rq";
    //        //dr = cmds1.ExecuteReader();
    //        //if (dr.Read())
    //        //{
    //        //    累计自产焦 = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
    //        //    累计落地焦 = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
    //        //}
    //        //dr.Close();
    //        for (int i = 0; i < 5; i++)
    //        {
    //            自产湿焦[5] += 自产湿焦[i];
    //            落地湿焦[5] += 落地湿焦[i];
    //            入炉焦炭[i] = 自产湿焦[i] * 0.96 * (1 - 自产焦水分 / 100) + 落地湿焦[i] * 0.96 * (1 - 落地焦水分 / 100);
    //            入炉焦炭[5] += 入炉焦炭[i];

    //            累计自产湿焦[5] += 累计自产湿焦[i];
    //            累计落地湿焦[5] += 累计落地湿焦[i];
    //            累计入炉焦炭[i] = 累计自产湿焦[i] * 0.96 * (1 - 累计自产焦平均水分 / 100) + 落地湿焦[i] * 0.96 * (1 - 累计落地焦平均水分 / 100);
    //            累计入炉焦炭[5] += 累计入炉焦炭[i];
    //        }

    //        全铁产量[5] = 全铁产量[0] + 全铁产量[1] + 全铁产量[2] + 全铁产量[3] + 全铁产量[4];
    //        SI小于07[5] = SI小于07[0] + SI小于07[1] + SI小于07[2] + SI小于07[3] + SI小于07[4];
    //        送炼钢铁[5] = 送炼钢铁[0] + 送炼钢铁[1] + 送炼钢铁[2] + 送炼钢铁[3] + 送炼钢铁[4];
    //        铸造铁[5] = 铸造铁[0] + 铸造铁[1] + 铸造铁[2] + 铸造铁[3] + 铸造铁[4];
    //        机烧[5] = 机烧[0] + 机烧[1] + 机烧[2] + 机烧[3] + 机烧[4];
    //        球团[5] = 球团[0] + 球团[1] + 球团[2] + 球团[3] + 球团[4];
    //        本溪[5] = 本溪[0] + 本溪[1] + 本溪[2] + 本溪[3] + 本溪[4];
    //        生料[5] = 生料[0] + 生料[1] + 生料[2] + 生料[3] + 生料[4];
    //       // 工艺称焦炭[5] = 工艺称焦炭[0] + 工艺称焦炭[1] + 工艺称焦炭[2] + 工艺称焦炭[3] + 工艺称焦炭[4];
    //        煤粉[5] = 煤粉[0] + 煤粉[1] + 煤粉[2] + 煤粉[3] + 煤粉[4];
                      
    //        累计全铁产量[5] = 累计全铁产量[0] + 累计全铁产量[1] + 累计全铁产量[2] + 累计全铁产量[3] + 累计全铁产量[4];
    //        累计SI小于07[5] = 累计SI小于07[0] + 累计SI小于07[1] + 累计SI小于07[2] + 累计SI小于07[3] + 累计SI小于07[4];
    //        累计送炼钢铁[5] = 累计送炼钢铁[0] + 累计送炼钢铁[1] + 累计送炼钢铁[2] + 累计送炼钢铁[3] + 累计送炼钢铁[4];
    //        累计铸造铁[5] = 累计铸造铁[0] + 累计铸造铁[1] + 累计铸造铁[2] + 累计铸造铁[3] + 累计铸造铁[4];
    //        累计机烧[5] = 累计机烧[0] + 累计机烧[1] + 累计机烧[2] + 累计机烧[3] + 累计机烧[4];
    //        累计球团[5] = 累计球团[0] + 累计球团[1] + 累计球团[2] + 累计球团[3] + 累计球团[4];
    //        累计本溪[5] = 累计本溪[0] + 累计本溪[1] + 累计本溪[2] + 累计本溪[3] + 累计本溪[4];
    //        累计生料[5] = 累计生料[0] + 累计生料[1] + 累计生料[2] + 累计生料[3] + 累计生料[4];
    //       // 累计工艺称焦炭[5] = 累计工艺称焦炭[0] + 累计工艺称焦炭[1] + 累计工艺称焦炭[2] + 累计工艺称焦炭[3] + 累计工艺称焦炭[4];
    //        累计煤粉[5] = 累计煤粉[0] + 累计煤粉[1] + 累计煤粉[2] + 累计煤粉[3] + 累计煤粉[4];

    //        for (int i = 0; i <6; i++)
    //        {
    //            if (全铁产量[i] > 0)
    //            {
    //                焦比[i] = 入炉焦炭[i] * 1000 / 全铁产量[i];
                   
    //            }
    //            if (全铁产量[i] > 0)
    //               煤比[i] = 煤粉[i] * 1000 / 全铁产量[i];
    //            if (累计全铁产量[i] > 0)
    //            {
    //                累计焦比[i] = 累计入炉焦炭[i] * 1000 / 累计全铁产量[i];
    //            } 
    //            if (累计全铁产量[i]>0)
    //                 累计煤比[i] = 累计煤粉[i] * 1000 / 累计全铁产量[i];
    //        }

    //        cn.Close();

    //    }

        
    //}


}
