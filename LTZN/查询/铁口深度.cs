using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;

namespace LTZN.查询
{
    class 铁口深度
    {
        public double[,] 情况1 = new double[5, 4];
        public double[,] 情况2 = new double[5, 4];
        public string[,] 情况1明细 = new string[5, 4];
        public string[,] 情况2明细 = new string[5, 4];

        public void display(HtmlDocument doc)
        {

            if (doc != null)
            {
                doc.OpenNew(true);
                doc.Write("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\" \"http://www.w3.org/TR/html4/loose.dtd\">");
                doc.Write("<HTML><head>");
                doc.Write("<script language=\"JavaScript\">  \n");
                doc.Write("function getMingXi(a,b,c)\n");
                doc.Write("{\n");
                doc.Write("var mx1=[");
                doc.Write("[\"" + 情况1明细[0, 0] + "\",\"" + 情况1明细[0, 1] + "\",\"" + 情况1明细[0, 2] + "\",\"" + 情况1明细[0, 3] + "\"],");
                doc.Write("[\"" + 情况1明细[1, 0] + "\",\"" + 情况1明细[1, 1] + "\",\"" + 情况1明细[1, 2] + "\",\"" + 情况1明细[1, 3] + "\"],");
                doc.Write("[\"" + 情况1明细[2, 0] + "\",\"" + 情况1明细[2, 1] + "\",\"" + 情况1明细[2, 2] + "\",\"" + 情况1明细[2, 3] + "\"],");
                doc.Write("[\"" + 情况1明细[3, 0] + "\",\"" + 情况1明细[3, 1] + "\",\"" + 情况1明细[3, 2] + "\",\"" + 情况1明细[3, 3] + "\"],");
                doc.Write("[\"" + 情况1明细[4, 0] + "\",\"" + 情况1明细[4, 1] + "\",\"" + 情况1明细[4, 2] + "\",\"" + 情况1明细[4, 3] + "\"]];\n");
                doc.Write("var mx2=[");
                doc.Write("[\"" + 情况2明细[0, 0] + "\",\"" + 情况2明细[0, 1] + "\",\"" + 情况2明细[0, 2] + "\",\"" + 情况2明细[0, 3] + "\"],");
                doc.Write("[\"" + 情况2明细[1, 0] + "\",\"" + 情况2明细[1, 1] + "\",\"" + 情况2明细[1, 2] + "\",\"" + 情况2明细[1, 3] + "\"],");
                doc.Write("[\"" + 情况2明细[2, 0] + "\",\"" + 情况2明细[2, 1] + "\",\"" + 情况2明细[2, 2] + "\",\"" + 情况2明细[2, 3] + "\"],");
                doc.Write("[\"" + 情况2明细[3, 0] + "\",\"" + 情况2明细[3, 1] + "\",\"" + 情况2明细[3, 2] + "\",\"" + 情况2明细[3, 3] + "\"],");
                doc.Write("[\"" + 情况2明细[4, 0] + "\",\"" + 情况2明细[4, 1] + "\",\"" + 情况2明细[4, 2] + "\",\"" + 情况2明细[4, 3] + "\"]];\n");
                doc.Write("if (c==1)");
                doc.Write("mingxi.innerHTML=mx1[a][b];\n");
                doc.Write("else\n");
                doc.Write("mingxi.innerHTML=mx2[a][b];\n");
                doc.Write("return;\n");
                doc.Write("}\n");
                doc.Write("</script>\n");
                doc.Write("</head><body bgcolor=#EEEEFF topmargin=25 leftmargin=5>");
                doc.Write("<h1 align=center>铁口深度统计</h1>");
                doc.Write("<hr width=80%>");
                doc.Write("<table  border=0 align=center><tr><td>");
                doc.Write("<table border=0 cellpadding=1 bgcolor=#cccccc  align=center>");
                doc.Write("<tr><td bgcolor=#C6FFC6 align=center height=30 width=56>高炉</td>");
                doc.Write("<td bgcolor=#C6FFC6 align=center height=30 width=56>班别</td>");
                doc.Write("<td bgcolor=#C6FFC6 align=center height=30 width=56>情况1</td>");
                doc.Write("<td bgcolor=#C6FFC6 align=center height=30 width=56>情况2</td>");
                doc.Write("<td bgcolor=#C6FFC6 align=center height=30 width=60>扣分</td></tr>");
                doc.Write("</tr>");
                for (int i = 0; i < 5; i++)
                {
                    doc.Write("<tr>");
                    doc.Write("<td bgcolor=#D7D7FF height=24 align=center  rowspan=4>" + (i + 1) + "高炉</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >甲班</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(情况1[i, 0],i,0,1) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(情况2[i, 0], i, 0, 2) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml1(情况1[i, 0] + 2*情况2[i, 0]) + "</td>");
                    doc.Write("</tr>");
                    doc.Write("<tr>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >乙班</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(情况1[i, 1], i, 1, 1) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(情况2[i, 1], i, 1, 2) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml1(情况1[i, 1] + 2 * 情况2[i, 1]) + "</td>");
                    doc.Write("</tr>");
                    doc.Write("<tr>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >丙班</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(情况1[i, 2], i, 2, 1) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(情况2[i, 2], i, 2, 2) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml1(情况1[i, 2] + 2 * 情况2[i, 2]) + "</td>");
                    doc.Write("</tr>");
                    doc.Write("<tr>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >丁班</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(情况1[i, 3], i, 3, 1) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml(情况2[i, 3], i, 3, 2) + "</td>");
                    doc.Write("<td bgcolor=#FFFFD7 align=right height=24 >" + convertToHtml1(情况1[i, 3] + 2 * 情况2[i, 3]) + "</td>");
                    doc.Write("</tr>");
                }
                doc.Write("</table>");
                doc.Write("</td>");
                doc.Write("<td valign=top><span id=\"mingxi\">");

                doc.Write("</span></td>");
                doc.Write("</tr></table>");
                doc.Write("<hr width=80%>");
               // doc.Write("<table  border=0 align=center><tr><td>情况１：1#、2#、4#、5# &nbsp;&nbsp;1400mm≤铁口深度≤1500mm &nbsp;；3# &nbsp;&nbsp;1200mm≤铁口深度≤1300mm &nbsp;　扣1分 </td></tr>");
              //  doc.Write("<tr><td>情况２：1#、2#、4#、5# &nbsp;&nbsp;铁口深度&lt;1400mm &nbsp;；3# &nbsp;&nbsp;铁口深度&lt;1200mm &nbsp;　扣2分 </td></tr></table>");
                doc.Write("<table  border=0 align=center><tr><td>情况１：1400mm≤铁口深度≤1500mm &nbsp;&nbsp;　扣1分 </td></tr>");
                doc.Write("<tr><td>情况２：铁口深度&lt;1400mm &nbsp; &nbsp;　扣2分 </td></tr></table>");
              
                doc.Write("</BODY></HTML>");
            }
        }
        private string convertToHtml(double d,int a,int b,int qingkuang)
        {
            if (d > 0)
                return d.ToString("<span style=\"cursor:hand\"  onclick=\"getMingXi(" + a + "," + b + "," + qingkuang + ");\">#######0次</span>");
            else
                return "&nbsp;";
        }
        private string convertToHtml1(double d)
        {
            if (d > 0)
                return d.ToString("#######0分");
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
           // cmd.CommandText = "SELECT GAOLU,高炉排班表.班别,COUNT(*) FROM DDLUCI INNER JOIN  高炉排班表 ON TRUNC(DDLUCI.ZDSJ) = 高炉排班表.日期 AND DDLUCI.BANCI = 高炉排班表.班次 WHERE TRUNC(ZDSJ)>=:RQ1 and TRUNC(ZDSJ)<=:RQ2 AND DKSJ IS NOT NULL and banluci<>1 and ((shendu >=1400 and shendu<=1500 and gaolu<>3) or (shendu >=1200 and shendu<=1300 and gaolu=3)) group by GAOLU,高炉排班表.班别";
            cmd.CommandText = "SELECT GAOLU,GETBANCI1(zdsj),COUNT(*) FROM DDLUCI  WHERE TRUNC(ZDSJ)>=:RQ1 and TRUNC(ZDSJ)<=:RQ2 AND DKSJ IS NOT NULL and shendu >=1400 and shendu<=1500 group by GAOLU,GETBANCI1(zdsj)";
            cmd.Parameters.Add(":RQ1", OracleType.DateTime).Value = rq1;
            cmd.Parameters.Add(":RQ2", OracleType.DateTime).Value = rq2;
            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    string banbie = dr.IsDBNull(1) ? "" : dr.GetString(1);
                    switch (banbie)
                    {
                        case "甲班":
                            情况1[gaolu - 1,0] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                            break;
                        case "乙班":
                            情况1[gaolu - 1, 1] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                            break;
                        case "丙班":
                            情况1[gaolu - 1, 2] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                            break;
                        case "丁班":
                            情况1[gaolu - 1, 3] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                            break;

                    }  
                }
            }
            dr.Close();
           // cmd.CommandText = "SELECT GAOLU,高炉排班表.班别,COUNT(*) FROM DDLUCI INNER JOIN  高炉排班表 ON TRUNC(DDLUCI.ZDSJ) = 高炉排班表.日期 AND DDLUCI.BANCI = 高炉排班表.班次 WHERE TRUNC(ZDSJ)>=:RQ1 and TRUNC(ZDSJ)<=:RQ2 AND DKSJ IS NOT NULL and banluci<>1 and ((shendu <1400 and gaolu<>3) or (shendu <1200 and gaolu=3)) group by GAOLU,高炉排班表.班别";
            cmd.CommandText = "SELECT GAOLU,GETBANCI1(zdsj),COUNT(*) FROM DDLUCI WHERE TRUNC(ZDSJ)>=:RQ1 and TRUNC(ZDSJ)<=:RQ2 AND DKSJ IS NOT NULL and shendu <1400  group by GAOLU,GETBANCI1(zdsj)";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    string banbie = dr.IsDBNull(1) ? "" : dr.GetString(1);
                    switch (banbie)
                    {
                        case "甲班":
                            情况2[gaolu - 1, 0] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                            break;
                        case "乙班":
                            情况2[gaolu - 1, 1] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                            break;
                        case "丙班":
                            情况2[gaolu - 1, 2] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                            break;
                        case "丁班":
                            情况2[gaolu - 1, 3] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                            break;

                    }
                }
            }
            dr.Close();

            string s = "<table border=0 cellpadding=1 bgcolor=#cccccc  align=center>";
            s += "<tr><td bgcolor=#C6FFC6 align=center height=30 width=66>日期</td>";
            s += "<td bgcolor=#C6FFC6 align=center height=30 width=56>班次</td>";
            s += "<td bgcolor=#C6FFC6 align=center height=30 width=66>本班炉次</td>";
            s += "<td bgcolor=#C6FFC6 align=center height=30 width=66>总炉次</td>";
            s += "<td bgcolor=#C6FFC6 align=center height=30 width=66>深度</td></tr>";

            for (int i = 0; i < 5; i++)
                for (int j = 0; j < 4; j++)
                {
                    情况1明细[i, j] = s;
                    情况2明细[i, j] = s;
                }


            //cmd.CommandText = "select * from (SELECT GAOLU,高炉排班表.班别 as 班别,ZDSJ,BANCI,BANLUCI,LUCI,SHENDU FROM DDLUCI INNER JOIN  高炉排班表 ON TRUNC(DDLUCI.ZDSJ) = 高炉排班表.日期 AND DDLUCI.BANCI = 高炉排班表.班次 WHERE TRUNC(ZDSJ)>=:RQ1 and TRUNC(ZDSJ)<=:RQ2 AND DKSJ IS NOT NULL and banluci<>1 and ((shendu >=1400 and shendu<=1500 and gaolu<>3) or (shendu >=1200 and shendu<=1300 and gaolu=3)) ";
            //cmd.CommandText += " UNION SELECT GAOLU,高炉排班上一班.班别 as 班别,ZDSJ,BANCI,BANLUCI,LUCI,SHENDU FROM DDLUCI INNER JOIN  高炉排班上一班 ON TRUNC(DDLUCI.ZDSJ) = 高炉排班上一班.日期 AND DDLUCI.BANCI = 高炉排班上一班.班次 WHERE TRUNC(ZDSJ)>=:RQ1 and TRUNC(ZDSJ)<=:RQ2 AND DKSJ IS NOT NULL and banluci=1 and ((shendu >=1400 and shendu<=1500 and gaolu<>3) or (shendu >=1200 and shendu<=1300 and gaolu=3))) order by GAOLU,班别,ZDSJ";
            cmd.CommandText = "SELECT GAOLU,GETBANCI1(zdsj) as 班别,ZDSJ,BANCI,BANLUCI,LUCI,SHENDU FROM DDLUCI WHERE TRUNC(ZDSJ)>=:RQ1 and TRUNC(ZDSJ)<=:RQ2 AND DKSJ IS NOT NULL and shendu >=1400 and shendu<=1500";
             
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0) && !dr.IsDBNull(1))
                {
                    int gaolu = dr.GetInt32(0);
                    int qingkuang = 4;
                    string banbie = dr.IsDBNull(1) ? "" : dr.GetString(1);
                    switch (banbie)
                    {
                        case "甲班":
                            qingkuang=0;
                            break;
                        case "乙班":
                            qingkuang=1;
                            break;
                        case "丙班":
                            qingkuang=2;
                            break;
                        case "丁班":
                            qingkuang=3;
                            break;

                    }
                    if (qingkuang < 4)
                        情况1明细[gaolu - 1, qingkuang] += "<tr><td bgcolor=#FFFFD7>" + (dr.IsDBNull(2) ? "" : dr.GetDateTime(2).ToLongDateString()) + "</td><td bgcolor=#FFFFD7>" + (dr.IsDBNull(3) ? "" : dr.GetString(3)) + "</td><td bgcolor=#FFFFD7>" + (dr.IsDBNull(4) ? "" : dr.GetInt32(4).ToString()) + "</td><td bgcolor=#FFFFD7>" + (dr.IsDBNull(5) ? "" : dr.GetString(5)) + "</td><td bgcolor=#FFFFD7>" + (dr.IsDBNull(6) ? "" : dr.GetDouble(6).ToString()) + "</td></tr>";
                }
                

            }
            dr.Close();

            //cmd.CommandText = "select * from (SELECT GAOLU,高炉排班表.班别 as 班别,ZDSJ,BANCI,BANLUCI,LUCI,SHENDU FROM DDLUCI INNER JOIN  高炉排班表 ON TRUNC(DDLUCI.ZDSJ) = 高炉排班表.日期 AND DDLUCI.BANCI = 高炉排班表.班次 WHERE TRUNC(ZDSJ)>=:RQ1 and TRUNC(ZDSJ)<=:RQ2 AND DKSJ IS NOT NULL and banluci<>1 and ((shendu <1400 and gaolu<>3) or (shendu <1200  and gaolu=3)) ";
            //cmd.CommandText += " UNION SELECT GAOLU,高炉排班上一班.班别 as 班别,ZDSJ,BANCI,BANLUCI,LUCI,SHENDU FROM DDLUCI INNER JOIN  高炉排班上一班 ON TRUNC(DDLUCI.ZDSJ) = 高炉排班上一班.日期 AND DDLUCI.BANCI = 高炉排班上一班.班次 WHERE TRUNC(ZDSJ)>=:RQ1 and TRUNC(ZDSJ)<=:RQ2 AND DKSJ IS NOT NULL and banluci=1 and ((shendu <1400  and gaolu<>3) or (shendu <1200 and gaolu=3))) order by GAOLU,班别,ZDSJ";
            cmd.CommandText = "SELECT GAOLU,GETBANCI1(zdsj) as 班别,ZDSJ,BANCI,BANLUCI,LUCI,SHENDU FROM DDLUCI WHERE TRUNC(ZDSJ)>=:RQ1 and TRUNC(ZDSJ)<=:RQ2 AND DKSJ IS NOT NULL and shendu <1400 ";
              
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0) && !dr.IsDBNull(1))
                {
                    int gaolu = dr.GetInt32(0);
                    int qingkuang = 4;
                    string banbie = dr.IsDBNull(1) ? "" : dr.GetString(1);
                    switch (banbie)
                    {
                        case "甲班":
                            qingkuang = 0;
                            break;
                        case "乙班":
                            qingkuang = 1;
                            break;
                        case "丙班":
                            qingkuang = 2;
                            break;
                        case "丁班":
                            qingkuang = 3;
                            break;

                    }
                    if (qingkuang < 4)
                        情况2明细[gaolu - 1, qingkuang] += "<tr><td bgcolor=#FFFFD7>" + (dr.IsDBNull(2) ? "" : dr.GetDateTime(2).ToLongDateString()) + "</td><td bgcolor=#FFFFD7>" + (dr.IsDBNull(3) ? "" : dr.GetString(3)) + "</td><td bgcolor=#FFFFD7>" + (dr.IsDBNull(4) ? "" : dr.GetInt32(4).ToString()) + "</td><td bgcolor=#FFFFD7>" + (dr.IsDBNull(5) ? "" : dr.GetString(5)) + "</td><td bgcolor=#FFFFD7>" + (dr.IsDBNull(6) ? "" : dr.GetDouble(6).ToString()) + "</td></tr>";
                }


            }
            dr.Close();

            for (int i = 0; i < 5; i++)
                for (int j = 0; j < 4; j++)
                {
                    情况1明细[i, j] +="</table>";
                    情况2明细[i, j] += "</table>";
                }
            cn.Close();
        }
    }
}
