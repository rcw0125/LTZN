using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;

namespace LTZN.查询
{
    class 大宗原料
    {
        private double[,] data = new double[7, 3];
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
                doc.Write("<h1 align=center>大宗原料查询</h1>");
                doc.Write("<hr width=80%>");
                doc.Write("<table border=0 cellpadding=1 bgcolor=#cccccc height=147 align=center>");
                doc.Write("<tr>");
                doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=center height=30 width=66>单位</td>");
                doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=center height=30 width=150>落地湿焦</td>");
                doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=center height=30 width=180>落地干焦(扣4％水)</td>");
                doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=center height=30 width=150>煤粉</td>");
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
                    doc.Write("<td height=26 bgcolor=#D7D7FF>" + danwei + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=26>" + convertToHtml(data[i, 0]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=26>" + convertToHtml(data[i, 1]) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=26>" + convertToHtml(data[i, 2]) + "</td>");
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
            cmd.CommandText = "select  高炉,SUM(落地湿焦),SUM(落地湿焦*0.96),SUM(煤粉) from 原料消耗 where trunc(日期)>=:rq1 and trunc(日期)<=:rq2 group by 高炉";
            cmd.Parameters.Add(":rq1", OracleType.DateTime).Value = rq1;
            cmd.Parameters.Add(":rq2", OracleType.DateTime).Value = rq2;
            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    data[gaolu - 1,0] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    data[gaolu - 1,1] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                    data[gaolu - 1,2] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                }
            }
            dr.Close();
            cn.Close();
            data[6, 0] = data[0, 0] + data[1, 0] + data[2, 0] + data[3, 0] + data[4, 0] + data[5, 0];
            data[6, 1] = data[0, 1] + data[1, 1] + data[2, 1] + data[3, 1] + data[4, 1] + data[5, 0];
            data[6, 2] = data[0, 2] + data[1, 2] + data[2, 2] + data[3, 2] + data[4, 2] + data[5, 0];

        }
    }
}
