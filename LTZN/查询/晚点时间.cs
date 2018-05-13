using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;

namespace LTZN.查询
{
    class 晚点时间
    {
        private double[] data = new double[7];
        private string[] gaoluTable = new string[7] { "", "", "", "", "", "", "" };

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
                      doc.Write("<html>\n");
                      doc.Write("<head>\n");
                      doc.Write("<meta http-equiv=\"Content-Type\" content=\"text/html; charset=gb2312\">\n");
                      doc.Write("<meta name=\"GENERATOR\" content=\"Microsoft FrontPage 4.0\">\n");
                      doc.Write("<meta name=\"ProgId\" content=\"FrontPage.Editor.Document\">\n");
                      doc.Write("<title>晚点时间查询</title>\n");
                      doc.Write("<style>\n");
                      doc.Write("<!--\n");
                      doc.Write("td           { background-color: #FFFFD7; text-align: center }\n");
                      doc.Write(".biaotou     { background-color: #C6FFC6; text-align: center }\n");
                      doc.Write(".td1     { background-color: #D7D7FF; text-align: center }\n");
                      doc.Write(".td2     { background-color: #EEEEFF; text-align: center }\n");
                      doc.Write("-->\n");
                      doc.Write("</style>\n");
                      doc.Write("<script language=\"JavaScript\">  \n");
                      doc.Write("function gaolu1()\n");
                      doc.Write("{\n");
                      if (gaoluTable[0] != "")
                      {
                          doc.Write("  if (g1.style.display==\"none\")\n");
                          doc.Write("        g1.style.display=\"\";\n");
                          doc.Write("   else\n");
                          doc.Write("        g1.style.display=\"none\";\n");
                          doc.Write("    \n");
                          doc.Write("   return;\n");
                      }
                      doc.Write("}\n");
                      doc.Write("function gaolu2()\n");
                      doc.Write("{\n");
                      if (gaoluTable[1] != "")
                      {
                          doc.Write("  if (g2.style.display==\"none\")\n");
                          doc.Write("        g2.style.display=\"\";\n");
                          doc.Write("   else\n");
                          doc.Write("        g2.style.display=\"none\";\n");
                          doc.Write("    \n");
                          doc.Write("   return;\n");
                      }
                      doc.Write("}\n");
                      doc.Write("function gaolu3()\n");
                      doc.Write("{\n");
                      if (gaoluTable[2] != "")
                      {
                          doc.Write("  if (g3.style.display==\"none\")\n");
                          doc.Write("        g3.style.display=\"\";\n");
                          doc.Write("   else\n");
                          doc.Write("        g3.style.display=\"none\";\n");
                          doc.Write("    \n");
                          doc.Write("   return;\n");
                      }
                      doc.Write("}\n");
                      doc.Write("function gaolu4()\n");
                      doc.Write("{\n");
                      if (gaoluTable[3] != "")
                      {
                          doc.Write("  if (g4.style.display==\"none\")\n");
                          doc.Write("        g4.style.display=\"\";\n");
                          doc.Write("   else\n");
                          doc.Write("        g4.style.display=\"none\";\n");
                          doc.Write("    \n");
                          doc.Write("   return;\n");
                      }
                      doc.Write("}\n");
                      doc.Write("function gaolu5()\n");
                      doc.Write("{\n");
                      if (gaoluTable[4] != "")
                      {
                          doc.Write("  if (g5.style.display==\"none\")\n");
                          doc.Write("        g5.style.display=\"\";\n");
                          doc.Write("   else\n");
                          doc.Write("        g5.style.display=\"none\";\n");
                          doc.Write("    \n");
                          doc.Write("   return;\n");
                      }
                      doc.Write("}\n");
                      doc.Write("function gaolu6()\n");
                      doc.Write("{\n");
                      if (gaoluTable[5] != "")
                      {
                          doc.Write("  if (g6.style.display==\"none\")\n");
                          doc.Write("        g6.style.display=\"\";\n");
                          doc.Write("   else\n");
                          doc.Write("        g6.style.display=\"none\";\n");
                          doc.Write("    \n");
                          doc.Write("   return;\n");
                      }
                      doc.Write("}\n");
                      doc.Write("</script>\n");
                      doc.Write("</head>\n");

                      doc.Write("<body bgcolor=\"#EEEEFF\" topmargin=\"25\" leftmargin=\"5\">\n");
                      doc.Write("<h1 align=\"center\">晚点时间查询</h1>\n");
                      doc.Write("<hr width=\"80%\">\n");
                      doc.Write("<table border=\"0\" cellpadding=\"1\" bgcolor=\"#cccccc\"  align=\"center\" width=760>\n");
                      doc.Write("<tr><td class=\"td1\" width=\"40%\"><span style=\"cursor:hand\"  onclick=\"gaolu1();\">一高炉</a></td><td>"+convertToHtml(data[0],"######0")+"分钟</td></tr>\n");
                      doc.Write("<tr><td  colspan=\"2\" class=\"td2\">");
                      doc.Write(gaoluTable[0]);                  
                      doc.Write("</td></tr>\n");
                      doc.Write("<tr><td class=\"td1\" width=\"40%\"><span style=\"cursor:hand\"  onclick=\"gaolu2();\">二高炉</a></td><td>" + convertToHtml(data[1], "######0") + "分钟</td></tr>\n");
                      doc.Write("<tr><td  colspan=\"2\" class=\"td2\">");
                      doc.Write(gaoluTable[1]);  
                      doc.Write("</td></tr>\n");
                      doc.Write("<tr><td class=\"td1\" width=\"40%\"><span style=\"cursor:hand\"  onclick=\"gaolu3();\">三高炉</a></td><td>" + convertToHtml(data[2], "######0") + "分钟</td></tr>\n");
                      doc.Write("<tr><td  colspan=\"2\" class=\"td2\">");
                      doc.Write(gaoluTable[2]);  
                      doc.Write("</td></tr>\n");
                      doc.Write("<tr><td class=\"td1\" width=\"40%\"><span style=\"cursor:hand\"  onclick=\"gaolu4();\">四高炉</a></td><td>" + convertToHtml(data[3], "######0") + "分钟</td></tr>\n");
                      doc.Write("<tr><td  colspan=\"2\" class=\"td2\">");
                      doc.Write(gaoluTable[3]);  
                      doc.Write("</td></tr>\n");
                      doc.Write("<tr><td class=\"td1\" width=\"40%\"><span style=\"cursor:hand\"  onclick=\"gaolu5();\">五高炉</a></td><td>" + convertToHtml(data[4], "######0") + "分钟</td></tr>\n");
                      doc.Write("<tr><td  colspan=\"2\" class=\"td2\">");
                      doc.Write(gaoluTable[4]);  
                      doc.Write("</td></tr>\n");
                      doc.Write("<tr><td class=\"td1\" width=\"40%\"><span style=\"cursor:hand\"  onclick=\"gaolu6();\">六高炉</a></td><td>" + convertToHtml(data[5], "######0") + "分钟</td></tr>\n");
                      doc.Write("<tr><td  colspan=\"2\" class=\"td2\">");
                      doc.Write(gaoluTable[5]);
                      doc.Write("</td></tr>\n");
                      doc.Write("</table>\n");
                      doc.Write("</body>\n");
                      doc.Write("</html>\n");



            }
        }
        public void getData(DateTime rq1, DateTime rq2)
        { 
            OracleConnection cn = new OracleConnection();
            cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
            cn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = cn;
            cmd.CommandText = "select  GAOLU,SUM(wdsj) from DDLUCI where trunc(ZDSJ)>=:rq1 and trunc(ZDSJ)<=:rq2 group by GAOLU";
            cmd.Parameters.Add(":rq1", OracleType.DateTime).Value = rq1;
            cmd.Parameters.Add(":rq2", OracleType.DateTime).Value = rq2;
            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    data[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                }
            }
            dr.Close();
            for (int i = 1; i <= 6; i++)
            {
                cmd.CommandText = "select rownum,trunc(zdsj),banci,luci,feliang,banluci,quchu,dksj,wdsj,zdsj,dgsj,tzsj from ddluci where trunc(ZDSJ)>=:rq1 and trunc(ZDSJ)<=:rq2 and wdsj>0 and gaolu=" + i;
                dr = cmd.ExecuteReader();
                if (dr.HasRows)
                {
                    gaoluTable[i - 1] += "<table border=\"0\" id=\"g" + i + "\" cellpadding=\"1\" bgcolor=\"#cccccc\" align=\"center\" style=\"display:none\">\n";
                    gaoluTable[i - 1] += "  <tr>\n";
                    gaoluTable[i - 1] += "    <td class=\"biaotou\" height=\"24\" width=\"36\">序号</td>\n";
                    gaoluTable[i - 1] += "    <td class=\"biaotou\" height=\"24\" width=\"80\">日期</td>\n";
                    gaoluTable[i - 1] += "    <td class=\"biaotou\" height=\"24\" width=\"40\">班次</td>\n";
                    gaoluTable[i - 1] += "    <td class=\"biaotou\" height=\"24\" width=\"70\">总炉次</td>\n";
                    gaoluTable[i - 1] += "    <td class=\"biaotou\" height=\"24\" width=\"66\">产量</td>\n";
                    gaoluTable[i - 1] += "    <td class=\"biaotou\" height=\"24\" width=\"56\">本班<br>炉次</td>\n";
                    gaoluTable[i - 1] += "    <td class=\"biaotou\" height=\"24\" width=\"56\">去向</td>\n";
                    gaoluTable[i - 1] += "    <td class=\"biaotou\" height=\"24\" width=\"56\">出铁<br>时间</td>\n";
                    gaoluTable[i - 1] += "    <td class=\"biaotou\" height=\"24\" width=\"56\">晚点<br>时间</td>\n";
                    gaoluTable[i - 1] += "    <td class=\"biaotou\" height=\"24\" width=\"56\">正点<br>时间</td>\n";
                    gaoluTable[i - 1] += "    <td class=\"biaotou\" height=\"24\" width=\"56\">回罐<br>时间</td>\n";
                    gaoluTable[i - 1] += "    <td class=\"biaotou\" height=\"24\" width=\"56\">通知<br>时间</td>\n";
                    gaoluTable[i - 1] += "  </tr>\n";
                    while (dr.Read())
                    {
                        gaoluTable[i - 1] += "<tr>\n";
                        gaoluTable[i - 1] += "<td height=24>" + (dr.IsDBNull(0) ? "&nbsp;" : dr.GetDouble(0).ToString()) + "</td>\n";
                        gaoluTable[i - 1] += "<td height=24>" + (dr.IsDBNull(1) ? "&nbsp;" : dr.GetDateTime(1).ToString("yyyy-MM-dd")) + "</td>\n";
                        gaoluTable[i - 1] += "<td height=24>" + (dr.IsDBNull(2) ? "&nbsp;" : dr.GetString(2)) + "</td>\n";
                        gaoluTable[i - 1] += "<td height=24>" + (dr.IsDBNull(3) ? "&nbsp;" : dr.GetString(3)) + "</td>\n";
                        gaoluTable[i - 1] += "<td height=24>" + (dr.IsDBNull(4) ? "&nbsp;" : dr.GetDouble(4).ToString()) + "</td>\n";
                        gaoluTable[i - 1] += "<td height=24>" + (dr.IsDBNull(5) ? "&nbsp;" : dr.GetInt32(5).ToString()) + "</td>\n";
                        gaoluTable[i - 1] += "<td height=24>" + (dr.IsDBNull(6) ? "&nbsp;" : dr.GetString(6)) + "</td>\n";
                        gaoluTable[i - 1] += "<td height=24>" + (dr.IsDBNull(7) ? "&nbsp;" : dr.GetDateTime(7).ToString("HH:MM")) + "</td>\n";
                        gaoluTable[i - 1] += "<td height=24>" + (dr.IsDBNull(8) ? "&nbsp;" : dr.GetDouble(8).ToString()) + "</td>\n";
                        gaoluTable[i - 1] += "<td height=24>" + (dr.IsDBNull(9) ? "&nbsp;" : dr.GetDateTime(9).ToString("HH:MM")) + "</td>\n";
                        gaoluTable[i - 1] += "<td height=24>" + (dr.IsDBNull(10) ? "&nbsp;" : dr.GetDateTime(10).ToString("HH:MM")) + "</td>\n";
                        gaoluTable[i - 1] += "<td height=24>" + (dr.IsDBNull(11) ? "&nbsp;" : dr.GetDateTime(11).ToString("HH:MM")) + "</td>\n";
                        gaoluTable[i - 1] += "</tr>\n";
                    }
                    dr.Close();
                    gaoluTable[i - 1] += "</table>\n";
                }
            }
            cn.Close();
        }
    }
}
