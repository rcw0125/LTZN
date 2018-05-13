using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;

namespace LTZN.查询
{
    class 机烧值比较
    {

        private double 二号皮带数 = 0;
        private double 三号皮带数 = 0;
        private double 总返矿量 = 0;

        private double 工艺称机烧 = 0;
        private double 落地焦炭 = 0;

        private double 生铁品位 = 0;
        private double 实际机烧值 = 0;
        private double 酸性粉含量 = 0;
        private double 机烧测算值 = 0;
        private double 返矿机烧 = 0;
        private double 入炉综合品位 = 0;


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
                doc.Write("<h1 align=center>机烧值比较</h1>");
                doc.Write("<hr width=80%>");
                doc.Write("<table border=0 cellpadding=1 bgcolor=#cccccc  align=center>");
                doc.Write("<tr>");
                doc.Write("<td bgcolor=#D7FFD7 align=center height=30 width=90>机烧测算值</td>");
                doc.Write("<td bgcolor=#FFFFD7 align=center height=30 width=120>" + convertToHtml(机烧测算值) + "</td>");
                doc.Write("<td bgcolor=#D7FFD7 align=center height=30 width=90>２＃皮带数</td>");
                doc.Write("<td bgcolor=#FFFFD7 align=center height=30 width=120>" + convertToHtml(二号皮带数) + "</td>");
                doc.Write("<td bgcolor=#D7FFD7 align=center height=30 width=90>酸性粉含量</td>");
                doc.Write("<td bgcolor=#FFFFD7 align=center height=30 width=120>" + convertToHtml(酸性粉含量) + "</td>");
                doc.Write("</tr>");
                doc.Write("<tr>");
                doc.Write("<td bgcolor=#D7FFD7 align=center height=30 >工艺称机烧</td>");
                doc.Write("<td bgcolor=#FFFFD7 align=center height=30 >" + convertToHtml(工艺称机烧) + "</td>");
                doc.Write("<td bgcolor=#D7FFD7 align=center height=30 >３＃称机烧</td>");
                doc.Write("<td bgcolor=#FFFFD7 align=center height=30 >" + convertToHtml(三号皮带数) + "</td>");
                doc.Write("<td bgcolor=#D7FFD7 align=center height=30 >落地焦炭</td>");
                doc.Write("<td bgcolor=#FFFFD7 align=center height=30 >" + convertToHtml(落地焦炭) + "</td>");
                doc.Write("</tr>");
                doc.Write("<tr>");
                doc.Write("<td bgcolor=#D7FFD7 align=center height=30 >实际机烧值</td>");
                doc.Write("<td bgcolor=#FFFFD7 align=center height=30 >" + convertToHtml(实际机烧值) + "</td>");
                doc.Write("<td bgcolor=#D7FFD7 align=center height=30 >总返矿量</td>");
                doc.Write("<td bgcolor=#FFFFD7 align=center height=30 >" + convertToHtml(总返矿量) + "</td>");
                doc.Write("<td bgcolor=#D7FFD7 align=center height=30 >生铁品位</td>");
                doc.Write("<td bgcolor=#FFFFD7 align=center height=30 >" + convertToHtml(生铁品位) + "</td>");
                doc.Write("</tr>");
                doc.Write("<tr>");
                doc.Write("<td bgcolor=#D7FFD7 align=center height=30 >&nbsp;</td>");
                doc.Write("<td bgcolor=#FFFFD7 align=center height=30 >&nbsp;</td>");
                doc.Write("<td bgcolor=#D7FFD7 align=center height=30 >返矿机烧</td>");
                doc.Write("<td bgcolor=#FFFFD7 align=center height=30 >" + convertToHtml(返矿机烧) + "</td>");
                doc.Write("<td bgcolor=#D7FFD7 align=center height=30 >入炉综合品位</td>");
                doc.Write("<td bgcolor=#FFFFD7 align=center height=30 >" + convertToHtml(入炉综合品位) + "</td>");
                doc.Write("</tr>");
                doc.Write("</table>");
                doc.Write("</BODY></HTML>");
            }
        }
        public void getData(DateTime rq1, DateTime rq2)
        {
            OracleConnection cn = new OracleConnection();
            cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
            cn.Open();

            OracleCommand cmd = new OracleCommand();
            cmd.Connection = cn;
            cmd.CommandText = "select sum(P05二号皮带),sum(P06三号皮带),sum(P07总返矿),sum(落地湿焦) from 全厂日消耗 where trunc(P01日期)>=:rq1 and trunc(P01日期)<=:rq2";
            cmd.Parameters.Add(":rq1", OracleType.DateTime).Value = rq1;
            cmd.Parameters.Add(":rq2", OracleType.DateTime).Value = rq2;

            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                二号皮带数 =dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
                三号皮带数 =dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                总返矿量=dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                落地焦炭 = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);

            }
            dr.Close();

            double 机烧=0, 竖炉球=0, 其他熟料=0, 本溪矿=0, 其他生料=0;
            double 机烧铁=0, 竖炉球铁=0, 其他熟料铁=0, 本溪矿铁=0, 其他生料铁=0;
            double 全铁产量=0;

            cmd.CommandText = "select SUM(机烧矿),SUM(竖球),SUM(熟料),SUM(本溪矿),SUM(生料) from 原料消耗 where trunc(日期)>=:rq1 and trunc(日期)<=:rq2";
            dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                机烧 = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
                竖炉球 = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                其他熟料 = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                本溪矿 = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                其他生料 = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
            }
            dr.Close();
            工艺称机烧 = 机烧;
            酸性粉含量 = (竖炉球 + 其他熟料 + 本溪矿 + 其他生料) * 8 / 92;
            实际机烧值 = 二号皮带数 + 三号皮带数 - 总返矿量 + 酸性粉含量;

            cmd.CommandText = "select sum(FELIANG),round(avg(100-NVL(FEC,0)-NVL(FESI,0)-NVL(FEMN,0)-NVL(FEP,0)-NVL(FES,0)-NVL(FETI,0)),2) from DDLUCI where  trunc(ZDSJ)>=:rq1 and trunc(ZDSJ)<=:rq2 and dksj is not null";
            dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                全铁产量 = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
                生铁品位 = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
            }
            dr.Close();


            cmd.CommandText = "select trunc(avg(P03TFE),2) from 日原料成分 where trunc(P01日期)>=:rq1 and trunc(P01日期)<=:rq2 and P02名称='机烧'";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                机烧铁 = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
            }
            dr.Close();
            cmd.CommandText = "select trunc(avg(P03TFE),2) from 日原料成分 where trunc(P01日期)>=:rq1 and trunc(P01日期)<=:rq2 and P02名称='竖炉球'";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                竖炉球铁 = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
            }
            dr.Close();
            cmd.CommandText = "select trunc(avg(P03TFE),2) from 日原料成分 where trunc(P01日期)>=:rq1 and trunc(P01日期)<=:rq2 and P02名称='本溪矿'";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                本溪矿铁 = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
            }
            dr.Close();
            cmd.CommandText = "select trunc(avg(P03TFE),2) from 日原料成分 where trunc(P01日期)>=:rq1 and trunc(P01日期)<=:rq2 and P02名称='其他生料'";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                其他生料铁 = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
            }
            dr.Close();

            cmd.CommandText = "select trunc(avg(P03TFE),2) from 日原料成分 where trunc(P01日期)>=:rq1 and trunc(P01日期)<=:rq2 and P02名称='其他熟料'";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                其他熟料铁 = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
            }
            dr.Close();
            机烧测算值 = (全铁产量 * 1.02 * 生铁品位 - (竖炉球 * 竖炉球铁 + 本溪矿 * 本溪矿铁 + 其他生料 * 其他生料铁 + 其他熟料 * 其他熟料铁)) / 机烧铁;


            //double 机烧品位 = 0;
            //double 竖炉球品位 = 0;

            //cmd.CommandText = "select trunc(AVG(TFE),2) from DDYL where mc<>'竖球' and trunc(SJ)>=:rq1 and trunc(SJ)<=:rq2";
            //dr = cmd.ExecuteReader();
            //if (dr.Read())
            //{

            //    机烧品位= dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
            //}
            //dr.Close();

            //cmd.CommandText = "select trunc(AVG(TFE),2) from DDYL where mc='竖球' and trunc(SJ)>=:rq1 and trunc(SJ)<=:rq2";
            //dr = cmd.ExecuteReader();
            //if (dr.Read())
            //{

            //      竖炉球品位= dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
            //}
            //dr.Close();

            //if(机烧+竖炉球>0)
            //    入炉综合品位 = (机烧*机烧品位+竖炉球*竖炉球品位)/(机烧+竖炉球);
            if (工艺称机烧 + 竖炉球 + 本溪矿 + 其他生料 + 其他熟料 > 0)
            {
                入炉综合品位 = (工艺称机烧 * 机烧铁 + 竖炉球 * 竖炉球铁 + 本溪矿 * 本溪矿铁 + 其他生料 * 其他生料铁 + 其他熟料 * 其他熟料铁) / (工艺称机烧 + 竖炉球 + 本溪矿 + 其他生料 + 其他熟料);
            }
             cn.Close();
        }
    }
}
