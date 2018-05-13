using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;

namespace LTZN.查询
{
    class 优质铁水率
    {
        private double[,] data = new double[7, 8];
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
                doc.Write("<h1 align=center>优质铁水率</h1>");
                doc.Write("<hr width=80%>");
                doc.Write("<table border=0 cellpadding=1 bgcolor=#cccccc  align=center>");
                doc.Write("<tr><td bgcolor=#C6FFC6 align=center height=24 width=56>单位</td>");
                doc.Write("<td bgcolor=#C6FFC6 align=center height=24 width=72>产量</td>");
                doc.Write("<td bgcolor=#C6FFC6 align=center height=24 width=72>优质铁</td>");
                doc.Write("<td bgcolor=#C6FFC6 align=center height=24 width=72>优质铁率(%)</td>");
                doc.Write("<td bgcolor=#C6FFC6 align=center height=24>碳<br>（加权平均）</td>");
                doc.Write("<td bgcolor=#C6FFC6 align=center height=24>硅<br>（加权平均）</td>");
                doc.Write("<td bgcolor=#C6FFC6 align=center height=24>锰<br>（加权平均）</td>");
                doc.Write("<td bgcolor=#C6FFC6 align=center height=24>磷<br>（加权平均）</td>");
                doc.Write("<td bgcolor=#C6FFC6 align=center height=24>硫<br>（加权平均）</td>");
                doc.Write("</tr>");
                for (int i = 0; i < 7; i++)
                {
                    string danwei = "";
                    switch (i)
                    {
                        case 0:
                            danwei = "一高炉";
                            break;
                        case 1:
                            danwei = "二高炉";
                            break;
                        case 2:
                            danwei = "三高炉";
                            break;
                        case 3:
                            danwei = "四高炉";
                            break;
                        case 4:
                            danwei = "五高炉";
                            break;
                        case 5:
                            danwei = "六高炉";
                            break;
                        case 6:
                            danwei = "全&nbsp;厂";
                            break;

                    }
                    doc.Write("<tr>");
                    doc.Write("<td height=24 bgcolor=#D7D7FF>" + danwei + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 0]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 1]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 2]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 3]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 4]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 5]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 6],"##0.000") + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 7], "##0.000") + "</td>");
                    doc.Write("</tr>");
                }
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
            cmd.CommandText = "SELECT GAOLU,SUM(FELIANG),TRUNC(SUM(FELIANG*FEC)/SUM(FELIANG),2),TRUNC(SUM(FELIANG*FESI)/SUM(FELIANG),2),TRUNC(SUM(FELIANG*FEMN)/SUM(FELIANG),2),TRUNC(SUM(FELIANG*FEP)/SUM(FELIANG),3),TRUNC(SUM(FELIANG*FES)/SUM(FELIANG),3) FROM DDLUCI WHERE TRUNC(ZDSJ)>=:RQ1 and TRUNC(ZDSJ)<=:RQ2 AND DKSJ IS NOT NULL GROUP BY GAOLU";
            cmd.Parameters.Add(":RQ1", OracleType.DateTime).Value = rq1;
            cmd.Parameters.Add(":RQ2", OracleType.DateTime).Value = rq2;
            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    data[gaolu - 1, 0] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    data[gaolu - 1, 3] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                    data[gaolu - 1, 4] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                    data[gaolu - 1, 5] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                    data[gaolu - 1, 6] = dr.IsDBNull(5) ? 0 : dr.GetDouble(5);
                    data[gaolu - 1, 7] = dr.IsDBNull(6) ? 0 : dr.GetDouble(6);
                }
            }
            dr.Close();
            cmd.CommandText = "SELECT GAOLU,SUM(FELIANG) FROM DDLUCI WHERE TRUNC(ZDSJ)>=:RQ1 and TRUNC(ZDSJ)<=:RQ2 AND DKSJ IS NOT NULL  AND ((ZDSJ<TO_DATE('2013-01-01','YYYY-MM-DD') and FESI<=0.70 AND FES<=0.04) OR (ZDSJ>=TO_DATE('2013-01-01','YYYY-MM-DD') and FESI<=0.70 AND FES<=0.04)) GROUP BY GAOLU";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    data[gaolu - 1, 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                }
            }
            dr.Close();
            cn.Close();
            data[6, 0] = data[0, 0] + data[1, 0] + data[2, 0] + data[3, 0] + data[4, 0] + data[5, 0];
            data[6, 1] = data[0, 1] + data[1, 1] + data[2, 1] + data[3, 1] + data[4, 1] + data[5, 1];
            for (int i = 0; i < 7; i++)
            {
                if (data[i, 0] > 0)
                    data[i, 2] = data[i, 1]*100 / data[i, 0];
                else
                    data[i, 2] = 0;
 
            }

            for (int i = 3; i < 8; i++)
            {
                if (data[6, 0] > 0)
                    data[6, i] = (data[0, 0] * data[0, i] + data[1, 0] * data[1, i] + data[2, 0] * data[2, i] + data[3, 0] * data[3, i] + data[4, 0] * data[4, i] + data[5, 0] * data[5, i]) / data[6, 0];
                else
                    data[6, i] = 0;
            }
        }
    }
}
