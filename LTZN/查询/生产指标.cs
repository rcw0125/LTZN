using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;

namespace LTZN.查询
{
    class 生产指标
    {
        private double[] data = new double[10];

        private string convertToHtml(double d)
        {
            if (d > 0)
                return d.ToString("#######0.00");
            else
                return "&nbsp;";
        }
        private string convertToHtml(double d, string format)
        {
            if (d > 0)
                return d.ToString(format);
            else
                return "&nbsp;";
        }
        public void display(HtmlDocument doc)
        {

            if (doc != null)
            {
                doc.OpenNew(true);
                doc.Write("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\" \"http://www.w3.org/TR/html4/loose.dtd\">");
                doc.Write("<HTML><body bgcolor=#EEEEFF topmargin=25 leftmargin=5>");
                doc.Write("<h1 align=center>生产指标</h1>");
                doc.Write("<hr width=80%>");
                doc.Write("<table border=0 cellpadding=1 bgcolor=#cccccc  align=center>");
                doc.Write("<tr>");
                doc.Write("<td bgcolor=#D7FFD7 align=center height=30 width=90>本日产量</td>");
                doc.Write("<td bgcolor=#FFFFD7 align=center height=30 width=150>" + convertToHtml(data[0]) + "</td>");
                doc.Write("<td bgcolor=#D7FFD7 align=center height=30 width=90>落地干焦</td>");
                doc.Write("<td bgcolor=#FFFFD7 align=center height=30 width=150>" + convertToHtml(data[1]) + "</td>");
                doc.Write("</tr>");
                doc.Write("<tr>");
                doc.Write("<td bgcolor=#D7FFD7 align=center height=30 width=90>湿煤粉</td>");
                doc.Write("<td bgcolor=#FFFFD7 align=center height=30 width=150>" + convertToHtml(data[2]) + "</td>");
                doc.Write("<td bgcolor=#D7FFD7 align=center height=30 width=90>差价含粉量</td>");
                doc.Write("<td bgcolor=#FFFFD7 align=center height=30 width=150>" + convertToHtml(data[3]) + "</td>");
                doc.Write("</tr>");
                doc.Write("<tr>");
                doc.Write("<td bgcolor=#D7FFD7 align=center height=30 width=90>干焦粉</td>");
                doc.Write("<td bgcolor=#FFFFD7 align=center height=30 width=150>" + convertToHtml(data[4]) + "</td>");
                doc.Write("<td bgcolor=#D7FFD7 align=center height=30 width=90>外购废铁</td>");
                doc.Write("<td bgcolor=#FFFFD7 align=center height=30 width=150>" + convertToHtml(data[5]) + "</td>");
                doc.Write("</tr>");
                doc.Write("<tr>");
                doc.Write("<td bgcolor=#D7FFD7 align=center height=30 width=90>富氧量</td>");
                doc.Write("<td bgcolor=#FFFFD7 align=center height=30 width=150>" + convertToHtml(data[6]) + "</td>");
                doc.Write("<td bgcolor=#D7FFD7 align=center height=30 width=90>机烧品位</td>");
                doc.Write("<td bgcolor=#FFFFD7 align=center height=30 width=150>" + convertToHtml(data[7]) + "</td>");
                doc.Write("</tr>");
                doc.Write("<tr>");
                doc.Write("<td bgcolor=#D7FFD7 align=center height=30 width=90>自产干焦</td>");
                doc.Write("<td bgcolor=#FFFFD7 align=center height=30 width=150>" + convertToHtml(data[8]) + "</td>");
                doc.Write("<td bgcolor=#D7FFD7 align=center height=30 width=90>机烧测算值</td>");
                doc.Write("<td bgcolor=#FFFFD7 align=center height=30 width=150>" + convertToHtml(data[9]) + "</td>");
                doc.Write("</tr>");
                doc.Write("</table>");
                doc.Write("</BODY></HTML>");
            }
        }
        public void getData(DateTime rq1, DateTime rq2)
        {
            double 全铁产量 = 0;
            double 湿焦粉 = 0;

            double 酸性含粉量 = 0;

            double 机烧 = 0;
            double 竖炉球 = 0;
            double 其他熟料 = 0;
            double 本溪矿 = 0;
            double 其他生料 = 0;

            double 机烧测算值 = 0;

            double 机烧铁 = 0;

            double 竖炉球铁 = 0;

            double 其他熟料铁 = 0;

            double 本溪矿铁 = 0;

            double 其他生料铁 = 0;

            double 铁品位 = 0;

            //double 落地湿焦 = 0;

            //double 自产湿焦 = 0;

            OracleConnection cn = new OracleConnection();
            cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
            cn.Open();
            String sqlstr = "select sum(FELIANG) from DDLUCI where trunc(ZDSJ)>=:rq1 and trunc(ZDSJ)<=:rq2 and dksj is not null";
            OracleCommand cmds1 = new OracleCommand(sqlstr, cn);
            cmds1.Parameters.Add(":rq1", OracleType.DateTime).Value = rq1;
            cmds1.Parameters.Add(":rq2", OracleType.DateTime).Value = rq2;
            OracleDataReader dr = cmds1.ExecuteReader();
            while (dr.Read())
            {
                data[0] = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
            }
            dr.Close();

            sqlstr = "select sum(feitie),sum(shijiaofen),sum(fuyang) from xiaohao where trunc(rq)>=:rq1 and trunc(rq)<=:rq2 ";
            cmds1 = new OracleCommand(sqlstr, cn);
            cmds1.Parameters.Add(":rq1", OracleType.DateTime).Value = rq1;
            cmds1.Parameters.Add(":rq2", OracleType.DateTime).Value = rq2;
            dr = cmds1.ExecuteReader();
            while (dr.Read())
            {

                data[5] = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
                湿焦粉 = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                data[6] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
            }
            dr.Close();
            sqlstr = "select sum(shijiaofen*(1-jiaofenshuifen/100)) from xiaohao where trunc(xiaohao.rq)>=:rq1 and trunc(xiaohao.rq)<=:rq2";
            cmds1 = new OracleCommand(sqlstr, cn);
            cmds1.Parameters.Add(":rq1", OracleType.DateTime).Value = rq1;
            cmds1.Parameters.Add(":rq2", OracleType.DateTime).Value = rq2;
            dr = cmds1.ExecuteReader();
            while (dr.Read())
            {
                data[4] = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);

            }
            dr.Close();

            sqlstr = "select SUM(GETZICHANG(P01日期)),SUM(GETLUODIG(P01日期)) from 全厂日消耗 where trunc(P01日期)>=:rq1 and trunc(P01日期)<=:rq2";
            cmds1 = new OracleCommand(sqlstr, cn);
            cmds1.Parameters.Add(":rq1", OracleType.DateTime).Value = rq1;
            cmds1.Parameters.Add(":rq2", OracleType.DateTime).Value = rq2;
            dr = cmds1.ExecuteReader();
            while (dr.Read())
            {

                data[8] = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
                data[1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
            }
            dr.Close();
            cmds1 = new OracleCommand("select SUM(机烧矿),SUM(竖球),SUM(熟料),SUM(本溪矿),SUM(生料),SUM(焦丁),SUM(煤粉) from 原料消耗 where trunc(日期)>=:rq1 and trunc(日期)<=:rq2", cn);
            cmds1.Parameters.Add(":rq1", OracleType.DateTime).Value = rq1.Date;
            cmds1.Parameters.Add(":rq2", OracleType.DateTime).Value = rq2.Date;
            dr = cmds1.ExecuteReader();
            if (dr.HasRows)
            {
                dr.Read();
                机烧 = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
                竖炉球 = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                其他熟料 = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                本溪矿 = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                其他生料 = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                酸性含粉量 = (竖炉球 + 其他熟料 + 本溪矿 + 其他生料) * 8 / 92;
               data[3] = Convert.ToDouble(酸性含粉量.ToString("##0.00"));
               data[2] = dr.IsDBNull(6) ? 0 : dr.GetDouble(6)/0.92;
               //data[8] = dr.IsDBNull(7) ? 0 : dr.GetDouble(7);
               //data[1] = dr.IsDBNull(8) ? 0 : dr.GetDouble(8);
            }
            dr.Close();
            sqlstr = "select sum(FELIANG),trunc(avg(100-NVL(FEC,0)-NVL(FESI,0)-NVL(FEMN,0)-NVL(FEP,0)-NVL(FES,0)-NVL(FETI,0)),2) from DDLUCI where  trunc(ZDSJ)>=:rq1 and trunc(ZDSJ)<=:rq2 and dksj is not null";
            cmds1 = new OracleCommand(sqlstr, cn);
            cmds1.Parameters.Add(":rq1", OracleType.DateTime).Value = rq1;
            cmds1.Parameters.Add(":rq2", OracleType.DateTime).Value = rq2;
            dr = cmds1.ExecuteReader();
            if (dr.Read())
            {
                全铁产量 = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
                铁品位 = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
            }
            dr.Close();
            sqlstr = "select trunc(avg(P03TFE),2) from 日原料成分 where trunc(P01日期)>=:rq1 and trunc(P01日期)<=:rq2 and P02名称='机烧'";
            cmds1 = new OracleCommand(sqlstr, cn);
            cmds1.Parameters.Add(":rq1", OracleType.DateTime).Value = rq1;
            cmds1.Parameters.Add(":rq2", OracleType.DateTime).Value = rq2;
            dr = cmds1.ExecuteReader();
            while (dr.Read())
            {
                data[7] = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
                机烧铁 = data[7];
            }
            dr.Close();
            sqlstr = "select trunc(avg(P03TFE),2) from 日原料成分 where trunc(P01日期)>=:rq1 and trunc(P01日期)<=:rq2 and P02名称='竖炉球'";
            cmds1 = new OracleCommand(sqlstr, cn);
            cmds1.Parameters.Add(":rq1", OracleType.DateTime).Value = rq1;
            cmds1.Parameters.Add(":rq2", OracleType.DateTime).Value = rq2;
            dr = cmds1.ExecuteReader();
            while (dr.Read())
            {
                竖炉球铁 = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
            }
            dr.Close();
            sqlstr = "select trunc(avg(P03TFE),2) from 日原料成分 where trunc(P01日期)>=:rq1 and trunc(P01日期)<=:rq2 and P02名称='本溪矿'";
            cmds1 = new OracleCommand(sqlstr, cn);
            cmds1.Parameters.Add(":rq1", OracleType.DateTime).Value = rq1;
            cmds1.Parameters.Add(":rq2", OracleType.DateTime).Value = rq2;
            dr = cmds1.ExecuteReader();
            while (dr.Read())
            {
                本溪矿铁 = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
            }
            dr.Close();
            sqlstr = "select trunc(avg(P03TFE),2) from 日原料成分 where trunc(P01日期)>=:rq1 and trunc(P01日期)<=:rq2 and P02名称='其他生料'";
            cmds1 = new OracleCommand(sqlstr, cn);
            cmds1.Parameters.Add(":rq1", OracleType.DateTime).Value = rq1;
            cmds1.Parameters.Add(":rq2", OracleType.DateTime).Value = rq2;
            dr = cmds1.ExecuteReader();
            while (dr.Read())
            {
                其他生料铁 = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
            }
            dr.Close();
            sqlstr = "select trunc(avg(P03TFE),2) from 日原料成分 where trunc(P01日期)>=:rq1 and trunc(P01日期)<=:rq2 and P02名称='其他熟料'";
            cmds1 = new OracleCommand(sqlstr, cn);
            cmds1.Parameters.Add(":rq1", OracleType.DateTime).Value = rq1;
            cmds1.Parameters.Add(":rq2", OracleType.DateTime).Value = rq2;
            dr = cmds1.ExecuteReader();
            while (dr.Read())
            {
                其他熟料铁 = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
            }
            dr.Close();

            if (机烧铁 > 0)
                机烧测算值 = (全铁产量 * 1.02 * 铁品位 - (竖炉球 * 竖炉球铁 + 本溪矿 * 本溪矿铁 + 其他生料 * 其他生料铁 + 其他熟料 * 其他熟料铁)) / 机烧铁;
            data[9] = Convert.ToDouble(机烧测算值.ToString("##0.00"));
            cn.Close();

        }
    }
}
