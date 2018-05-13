using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;
using LTZN.Repository;
using LTZN.CalFramework;

namespace LTZN.查询
{
    public class 其它数据 : QueryModel
    {
        public 其它数据()
            : base("其它数据")
        {

        }

        public void getData(DateTime rq)
        {
            this.myModel.Init();

            conn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;

            //各班产量
            cmd.CommandText = "select gaolu,banci,sum(feliang) as banci,sum(feliang) from ddluci where  trunc(zdsj)=:rq and dksj is not null group by gaolu,banci";
            cmd.Parameters.Add(":rq", OracleType.DateTime).Value = rq;
            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0) && !dr.IsDBNull(1) && !dr.IsDBNull(2))
                {
                    int gaolu = dr.GetInt32(0);
                    string banci = dr.GetString(1);
                    this.SetValue("高炉" + gaolu.ToString() + "." + banci, dr.GetDouble(2));
                }
            }
            dr.Close();

            //Si大小分类产量
            StringBuilder sqlSb=new StringBuilder();
            sqlSb.Append( "select gaolu,");
             sqlSb.Append( "case when fesi<=0.45  then 'L04' ");
              sqlSb.Append( "when  fesi<=0.85 then 'L08' ");
              sqlSb.Append( "when  fesi<=1.25 then 'L10' ");
              sqlSb.Append( "when  fesi<1.6 then 'Z14' ");
              sqlSb.Append( "when  fesi<=2 then 'Z18' ");
              sqlSb.Append( "when  fesi<=2.4 then 'Z22' ");
              sqlSb.Append( "ELSE 'Z26' ");
              sqlSb.Append( "end,");
              sqlSb.Append("sum(feliang) from ddluci where trunc(zdsj)=:rq and dksj is not null group by gaolu,");
              sqlSb.Append("case when fesi<=0.45  then 'L04' ");
              sqlSb.Append("when  fesi<=0.85 then 'L08' ");
              sqlSb.Append("when  fesi<=1.25 then 'L10' ");
              sqlSb.Append("when  fesi<1.6 then 'Z14' ");
              sqlSb.Append("when  fesi<=2 then 'Z18' ");
              sqlSb.Append("when  fesi<=2.4 then 'Z22' ");
              sqlSb.Append("else 'Z26' ");
              sqlSb.Append("end");
            cmd.CommandText = sqlSb.ToString();
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    if (!dr.IsDBNull(1))
                        this.SetValue(string.Format("高炉{0}.{1}", gaolu, dr.GetString(1)), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                }
            }
            dr.Close();

            //各高炉累计全铁产量
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0) && !dr.IsDBNull(1))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("高炉{0}.累计产量", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();

            cmd.CommandText = "select 高炉,机烧矿,球团矿, 国内球团矿,进口球团矿,PB块,纽曼块,其它块矿,高钛球团矿,高品位钛球,煤粉,焦丁,工艺称,富氧量,自产湿焦,落地湿焦,硅石,萤石,蛇纹石,其它熔剂 from 原料消耗 where 日期=:rq";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("高炉{0}.机烧", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                    this.SetValue(string.Format("高炉{0}.球团矿", gaolu), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                    this.SetValue(string.Format("高炉{0}.国内球团矿", gaolu), dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                    this.SetValue(string.Format("高炉{0}.进口球团矿", gaolu), dr.IsDBNull(4) ? 0 : dr.GetDouble(4));
                    this.SetValue(string.Format("高炉{0}.PB块", gaolu), dr.IsDBNull(5) ? 0 : dr.GetDouble(5));
                    this.SetValue(string.Format("高炉{0}.纽曼块", gaolu), dr.IsDBNull(6) ? 0 : dr.GetDouble(6));

                    this.SetValue(string.Format("高炉{0}.其它块矿", gaolu), dr.IsDBNull(7) ? 0 : dr.GetDouble(7));
                    this.SetValue(string.Format("高炉{0}.高钛球团矿", gaolu), dr.IsDBNull(8) ? 0 : dr.GetDouble(8));
                    this.SetValue(string.Format("高炉{0}.高品位钛球", gaolu), dr.IsDBNull(9) ? 0 : dr.GetDouble(9));
                    this.SetValue(string.Format("高炉{0}.煤粉", gaolu), dr.IsDBNull(10) ? 0 : dr.GetDouble(10));
                    this.SetValue(string.Format("高炉{0}.焦丁", gaolu), dr.IsDBNull(11) ? 0 : dr.GetDouble(11));
                    this.SetValue(string.Format("高炉{0}.工艺称湿焦", gaolu), dr.IsDBNull(12) ? 0 : dr.GetDouble(12));
                    this.SetValue(string.Format("高炉{0}.富氧量", gaolu), dr.IsDBNull(13) ? 0 : dr.GetDouble(13));
                    this.SetValue(string.Format("高炉{0}.自产湿焦", gaolu), dr.IsDBNull(14) ? 0 : dr.GetDouble(14));
                    this.SetValue(string.Format("高炉{0}.落地湿焦", gaolu), dr.IsDBNull(15) ? 0 : dr.GetDouble(15));
                    this.SetValue(string.Format("高炉{0}.硅石", gaolu), dr.IsDBNull(16) ? 0 : dr.GetDouble(16));
                    this.SetValue(string.Format("高炉{0}.萤石", gaolu), dr.IsDBNull(17) ? 0 : dr.GetDouble(17));
                    this.SetValue(string.Format("高炉{0}.蛇纹石", gaolu), dr.IsDBNull(18) ? 0 : dr.GetDouble(18));
                    this.SetValue(string.Format("高炉{0}.其它熔剂", gaolu), dr.IsDBNull(19) ? 0 : dr.GetDouble(19));
                }
            }
            dr.Close();

            cmd.CommandText = "select trunc(avg(shuif),2) from ddjt where trunc(sj)=:RQ and mc in ('自产焦夜','自产焦白','自产焦中')";
            dr = cmd.ExecuteReader();
            if (dr.Read())
                this.SetValue("自产湿焦水份", dr.IsDBNull(0) ? 0 : dr.GetDouble(0));
            dr.Close();
            cmd.CommandText = "select trunc(avg(shuif),2) from ddjt where trunc(sj)=:RQ and mc='外进焦'";
            dr = cmd.ExecuteReader();
            if (dr.Read())
                this.SetValue("落地湿焦水份", dr.IsDBNull(0) ? 0 : dr.GetDouble(0));
            dr.Close();

            //碎铁
            cmd.CommandText = "select gaolu,feitie from xiaohao where rq=:rq";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("高炉{0}.碎铁", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }

            dr.Close();

            //大中修
            cmd.CommandText = "select 高炉,sum(间隔) from 休风 where  trunc(时间)=:rq and (分类='大修' or 分类='中修')  and  间隔>0　and 间隔 is not null group by 高炉";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                int gaolu = dr.IsDBNull(0) ? 6 : dr.GetInt32(0);
                this.SetValue(string.Format("高炉{0}.大中修", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));

            }
            dr.Close();


            this.SetValue("累计天数", rq.Day);
            conn.Close();
        }
    
    }

    //class 其它数据
    //{
    //    private int gaoluShu = 6;

    //    private double[,] data = new double[7, 36];
    //    public double[] jfPrecent = new double[6];
    //    public double[] jfAmount = new double[6];
    //    public bool[] isPrecent = new bool[6];


    //    private string convertToHtml(double d)
    //    {
    //        if (d > 0)
    //            return "<div style=\"font-size:10t;\">" + d.ToString("#######0.00") + "</div>";
    //        else
    //            return "&nbsp;";
    //    }
    //    private string convertToHtml(double d, string format)
    //    {
    //        if (d > 0)
    //            return "<div style=\"font-size:10t;\">" + d.ToString(format) + "</div>";
    //        else
    //            return "&nbsp;";
    //    }
    //    public void display(HtmlDocument doc)
    //    {
            
    //        if (doc != null)
    //        {
    //                doc.OpenNew(true);
    //                doc.Write("<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.01 Transitional//EN\" \"http://www.w3.org/TR/html4/loose.dtd\">");
    //                doc.Write("<HTML><body bgcolor=#EEEEFF topmargin=25 leftmargin=5>");
    //                doc.Write("<h1 align=center>综合厂调指标查询</h1>");
    //                doc.Write("<hr width=80%>");
    //                doc.Write("<table border=0 cellpadding=1 bgcolor=#cccccc height=147 align=center>");
    //                doc.Write("<tr>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=56>单位</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=56>夜班</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=56>白班</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=56>中班</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=76>全铁产量</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=82>累计产量</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=76>平均产量</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=61>L04</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=61>L08</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=38>L10</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=38>Z14</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=38>Z18</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=38>Z22</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=38>Z26</td>");
    //                  doc.Write("</tr>");
    //                  for (int i = 0; i < gaoluShu+1; i++)
    //                  {
    //                      string danwei="";
    //                      switch (i)
    //                      {
    //                          case 0:
    //                              danwei="一高炉";
    //                              break;
    //                              case 1:
    //                              danwei="二高炉";
    //                              break;
    //                              case 2:
    //                              danwei="三高炉";
    //                              break;
    //                              case 3:
    //                              danwei="四高炉";
    //                              break;
    //                              case 4:
    //                              danwei="五高炉";
    //                              break;
    //                              case 5:
    //                              danwei = "六高炉";
    //                              break;
    //                              case 6:
    //                              danwei="全&nbsp;厂";
    //                              break;

    //                      }
    //                      doc.Write("<tr>");
    //                      doc.Write("<td height=24 bgcolor=#D7D7FF>"+danwei+"</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 0]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 1]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 2]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 3]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 4]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 5]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 6]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 7]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 8]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 9]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 10]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 11]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 12]) + "</td>");
    //                      doc.Write("</tr>");
    //                  }
    //                  doc.Write("</table>");
    //                  doc.Write("<table border=0 cellpadding=1 bgcolor=#cccccc height=147  align=center>");
    //                  doc.Write("<tr>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=56>单位</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=80>七号称湿焦</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=80>工艺称湿焦</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=67>矿耗总和</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=55>机烧</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=45>竖球</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=45>印球</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=45>熟料</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=49>本溪矿</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=45>生料</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=49>煤粉</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=45>碎铁</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=45>焦丁</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=66>焦炭总耗</td>");
    //                  doc.Write("</tr>");
    //                  for (int i = 0; i < gaoluShu + 1; i++)
    //                  {
    //                      string danwei = "";
    //                      switch (i)
    //                      {
    //                          case 0:
    //                              danwei = "一高炉";
    //                              break;
    //                          case 1:
    //                              danwei = "二高炉";
    //                              break;
    //                          case 2:
    //                              danwei = "三高炉";
    //                              break;
    //                          case 3:
    //                              danwei = "四高炉";
    //                              break;
    //                          case 4:
    //                              danwei = "五高炉";
    //                              break;
    //                          case 5:
    //                              danwei = "六高炉";
    //                              break;
    //                          case 6:
    //                              danwei = "全&nbsp;厂";
    //                              break;

    //                      }
    //                      doc.Write("<tr>");
    //                      doc.Write("<td height=24 bgcolor=#D7D7FF>" + danwei + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 13]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 14]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 15]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 16]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 17]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 18]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 19]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 20]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 21]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 22]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 23]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 24]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 25]) + "</td>");
    //                      doc.Write("</tr>");
    //                  }
    //                  doc.Write("</table>");
    //                  doc.Write("<table border=0 cellpadding=1 bgcolor=#cccccc height=147  align=center>");
    //                  doc.Write("<tr>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=56>单位</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=80>七号称焦比</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=78>综合焦比</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=80>工艺称焦比</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=70>实物系数</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=70>冶炼强度</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=70>焦炭负荷</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=70>碎铁比</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=70>煤比</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=70>回收率</td>");
    //                  doc.Write("<td bgcolor=#C6FFC6 bgcolor=#FFFFD7 align=right height=24 width=70>富氧率</td>");
    //                  doc.Write("</tr>");
    //                  for (int i = 0; i < gaoluShu + 1; i++)
    //                  {
    //                      string danwei="";
    //                      switch (i)
    //                      {
    //                          case 0:
    //                              danwei="一高炉";
    //                              break;
    //                              case 1:
    //                              danwei="二高炉";
    //                              break;
    //                              case 2:
    //                              danwei="三高炉";
    //                              break;
    //                              case 3:
    //                              danwei="四高炉";
    //                              break;
    //                              case 4:
    //                              danwei="五高炉";
    //                              break;
    //                              case 5:
    //                              danwei = "六高炉";
    //                              break;
    //                              case 6:
    //                              danwei="全&nbsp;厂";
    //                              break;

    //                      }
    //                      doc.Write("<tr>");
    //                      doc.Write("<td height=24 bgcolor=#D7D7FF>"+danwei+"</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 26],"#####0") + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 27], "#####0") + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 28], "#####0") + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 29]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 30]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 31]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 32], "#####0") + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 33], "#####0") + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 34]) + "</td>");
    //                      doc.Write("<td bgcolor=#FFFFD7 align=right height=24>" + convertToHtml(data[i, 35]) + "</td>");
    //                      doc.Write("</tr>");
    //                  }
    //                  doc.Write("</table>");
    //                  doc.Write("</BODY></HTML>");
    //                  Program.Log(doc.Body.InnerHtml); 
              
    //        }
    //    }

    //    public void loadParam()
    //    {
    //        Ltzn_ParaTable tm = new Ltzn_ParaTable();
    //        tm.LoadLikeName("焦粉%");

    //        for (int i = 0; i < gaoluShu; i++)
    //        {
    //            isPrecent[i] = true;
    //            jfPrecent[i] = 4;
    //        }
    //        foreach (var item in tm)
    //        {
    //            for (int i = 0; i < gaoluShu; i++)
    //            {
    //                if (item.NAME == "焦粉" + (i + 1).ToString() + "高炉比例")
    //                {
    //                    jfPrecent[i] = item.VAL;
    //                }
    //                if (item.NAME == "焦粉" + (i + 1).ToString() + "高炉重量")
    //                {
    //                    jfAmount[i] = item.VAL;
    //                }
    //                if (item.NAME == "焦粉" + (i + 1).ToString() + "高炉是否比例")
    //                {
    //                    isPrecent[i] = (item.VAL == 1);
    //                }
    //            }
    //        }
    //    }
    //    public void saveParam()
    //    {
    //        Ltzn_ParaTable.ClearTableLikeName("焦粉%");

    //        Ltzn_ParaTable tm = new Ltzn_ParaTable();

    //        for (int i = 0; i < gaoluShu; i++)
    //        {
    //            Ltzn_Para lp1 = new Ltzn_Para();
    //            lp1.NAME = "焦粉" + (i + 1).ToString() + "高炉是否比例";
    //            if (isPrecent[i])
    //                lp1.VAL = 1;
    //            else
    //                lp1.VAL = 0;

    //            Ltzn_Para lp2 = new Ltzn_Para();
    //            lp2.NAME = "焦粉" + (i + 1).ToString() + "高炉比例";
    //            lp2.VAL = jfPrecent[i];

    //            Ltzn_Para lp3 = new Ltzn_Para();
    //            lp3.NAME = "焦粉" + (i + 1).ToString() + "高炉重量";
    //            lp3.VAL = jfAmount[i];

    //            tm.Add(lp1);
    //            tm.Add(lp2);
    //            tm.Add(lp3);

    //        }

    //        tm.Save();
    //    }

    //    public void getData(DateTime rq)
    //    {
    //        OracleConnection cn = new OracleConnection();
    //        cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
    //        cn.Open();
    //        OracleCommand cmd = new OracleCommand();
    //        cmd.Connection = cn;

    //        //各班产量
    //        cmd.CommandText = "select gaolu,banci,sum(feliang) from ddluci where trunc(zdsj)=:rq and dksj is not null group by gaolu,banci";
    //        cmd.Parameters.Add(":rq", OracleType.DateTime).Value = rq;
    //        OracleDataReader dr = cmd.ExecuteReader();
    //        while (dr.Read())
    //        {
    //            if (!dr.IsDBNull(0) && !dr.IsDBNull(1) && !dr.IsDBNull(2))
    //            {
    //                int gaolu = dr.GetInt32(0);
    //                string banci = dr.GetString(1);
    //                switch (banci)
    //                {
    //                    case "夜班":
    //                        data[gaolu - 1, 0] = dr.GetDouble(2);
    //                        break;
    //                    case "白班":
    //                        data[gaolu - 1, 1] = dr.GetDouble(2);
    //                        break;
    //                    case "中班":
    //                        data[gaolu - 1, 2] = dr.GetDouble(2);
    //                        break;
    //                }
    //            }
    //        }
    //        dr.Close();
    //        //L04
    //        cmd.CommandText = "select gaolu,sum(feliang) from ddluci where fesi<=0.45  and  trunc(zdsj)=:rq and dksj is not null group by gaolu";
    //        dr = cmd.ExecuteReader();
    //        while (dr.Read())
    //        {
    //            if (!dr.IsDBNull(0))
    //            {
    //                int gaolu = dr.GetInt32(0);
    //                data[gaolu - 1, 6] += dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
    //            }
    //        }
    //        dr.Close();
    //        //L08
    //        cmd.CommandText = "select gaolu,sum(feliang) from ddluci where fesi>0.45 and fesi<=0.85  and  trunc(zdsj)=:rq and dksj is not null group by gaolu";
    //        dr = cmd.ExecuteReader();
    //        while (dr.Read())
    //        {
    //            if (!dr.IsDBNull(0))
    //            {
    //                int gaolu = dr.GetInt32(0);
    //                data[gaolu - 1, 7] += dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
    //            }
    //        }
    //        dr.Close();
    //        //L10
    //        cmd.CommandText = "select gaolu,sum(feliang) from ddluci where fesi>0.85 and fesi<=1.25  and  trunc(zdsj)=:rq and dksj is not null group by gaolu";
    //        dr = cmd.ExecuteReader();
    //        while (dr.Read())
    //        {
    //            if (!dr.IsDBNull(0))
    //            {
    //                int gaolu = dr.GetInt32(0);
    //                data[gaolu - 1, 8] += dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
    //            }
    //        }
    //        dr.Close();

    //        //Z14
    //        cmd.CommandText = "select gaolu,sum(feliang) from ddluci where fesi>1.25 and fesi<1.6  and  trunc(zdsj)=:rq and dksj is not null group by gaolu";
    //        dr = cmd.ExecuteReader();
    //        while (dr.Read())
    //        {
    //            if (!dr.IsDBNull(0))
    //            {
    //                int gaolu = dr.GetInt32(0);
    //                data[gaolu - 1, 9] += dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
    //            }
    //        }
    //        dr.Close();
    //        //Z18
    //        cmd.CommandText = "select gaolu,sum(feliang) from ddluci where fesi>=1.6 and fesi<=2.0  and  trunc(zdsj)=:rq and dksj is not null group by gaolu";
    //        dr = cmd.ExecuteReader();
    //        while (dr.Read())
    //        {
    //            if (!dr.IsDBNull(0))
    //            {
    //                int gaolu = dr.GetInt32(0);
    //                data[gaolu - 1, 10] += dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
    //            }
    //        }
    //        dr.Close();
    //        //Z22
    //        cmd.CommandText = "select gaolu,sum(feliang) from ddluci where fesi>2.0 and fesi<=2.4  and  trunc(zdsj)=:rq and dksj is not null group by gaolu";
    //        dr = cmd.ExecuteReader();
    //        while (dr.Read())
    //        {
    //            if (!dr.IsDBNull(0))
    //            {
    //                int gaolu = dr.GetInt32(0);
    //                data[gaolu - 1, 11] += dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
    //            }
    //        }
    //        dr.Close();
    //        //Z26以上
    //        cmd.CommandText = "select gaolu,sum(feliang) from ddluci where fesi>2.4  and  trunc(zdsj)=:rq and dksj is not null group by gaolu";
    //        dr = cmd.ExecuteReader();
    //        while (dr.Read())
    //        {
    //            if (!dr.IsDBNull(0))
    //            {
    //                int gaolu = dr.GetInt32(0);
    //                data[gaolu - 1, 12] += dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
    //            }
    //        }
    //        dr.Close();

    //        //各高炉全铁产量
    //        cmd.CommandText = "select gaolu,sum(feliang) from ddluci where trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and dksj is not null group by gaolu";
    //        dr = cmd.ExecuteReader();
    //        while (dr.Read())
    //        {
    //            if (!dr.IsDBNull(0) && !dr.IsDBNull(1))
    //            {
    //                int gaolu = dr.GetInt32(0);
    //                data[gaolu - 1, 4] = dr.GetDouble(1);
    //            }

    //        }
    //        dr.Close();

            
    //        double[] fy = new double[7]; //富氧量
    //        double[] lf = new double[7]; //冷风流量
    //        lf[0] = lf[1] = lf[3] = 84000;
    //        lf[2] = 54000;
    //        lf[4] = 108000;
    //        lf[5] = 108000;

    //        double[] zc = new double[7]; //自产湿焦
    //        double[] ld = new double[7]; //落地湿焦
    //        cmd.CommandText ="select 高炉,机烧矿,竖球,熟料,本溪矿,生料,煤粉,焦丁,工艺称,富氧量,自产湿焦,落地湿焦 from 原料消耗 where 日期=:rq";
    //        dr = cmd.ExecuteReader();
    //        while (dr.Read())
    //        {
    //            if (!dr.IsDBNull(0))
    //            {
    //                int gaolu = dr.GetInt32(0);
    //                data[gaolu - 1, 16] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
    //                data[gaolu - 1, 17] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
    //                data[gaolu - 1, 19] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
    //                data[gaolu - 1, 20] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
    //                data[gaolu - 1, 21] = dr.IsDBNull(5) ? 0 : dr.GetDouble(5);
    //                data[gaolu - 1, 22] = dr.IsDBNull(6) ? 0 : dr.GetDouble(6);
    //                data[gaolu - 1, 24] = dr.IsDBNull(7) ? 0 : dr.GetDouble(7);
    //                data[gaolu - 1, 14] = dr.IsDBNull(8) ? 0 : dr.GetDouble(8);
    //                fy[gaolu - 1] = dr.IsDBNull(9) ? 0 : dr.GetDouble(9);
    //                zc[gaolu - 1] = dr.IsDBNull(10) ? 0 : dr.GetDouble(10);
    //                ld[gaolu - 1] = dr.IsDBNull(11) ? 0 : dr.GetDouble(11);
    //                data[gaolu - 1, 13] = zc[gaolu - 1] + ld[gaolu - 1];
    //            }

    //        }
    //        dr.Close();

    //        double 自产焦水分 = 0;
    //        double 落地焦水分 = 0;
    //       // double 自产焦 = 0;
    //        //double 落地焦 = 0;

    //        cmd.CommandText = "select trunc(avg(shuif),2) from ddjt where trunc(sj)=:RQ and mc in ('自产焦夜','自产焦白','自产焦中')";
    //        dr = cmd.ExecuteReader();
    //        if (dr.Read())
    //            自产焦水分 = (dr.IsDBNull(0) ? 0 : dr.GetDouble(0));
    //        dr.Close();
    //        cmd.CommandText = "select trunc(avg(shuif),2) from ddjt where trunc(sj)=:RQ and mc='外进焦'";
    //        dr = cmd.ExecuteReader();
    //        if (dr.Read())
    //            落地焦水分 = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
    //        dr.Close();

    //        ////全厂七号称
    //        //cmd.CommandText = "select 自产湿焦,落地湿焦 from 全厂日消耗 where P01日期=:rq";
    //        //dr = cmd.ExecuteReader();
    //        //if (dr.Read())
    //        //{
    //        //    自产焦=dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
    //        //    落地焦= dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
    //        //}
    //        //dr.Close();
    //        data[6, 13] = data[0, 13]+data[1, 13] + data[2, 13] + data[3, 13] + data[4, 13] + data[5, 13];               //七号称
    //        zc[6] = zc[0] + zc[1] + zc[2] + zc[3] + zc[4] + zc[5];
    //        ld[6] = ld[0] + ld[1] + ld[2] + ld[3] + ld[4] + ld[5];

    //        data[6, 25] = 0;
    //        for (int i = 0; i < gaoluShu; i++)
    //        {
    //            if (isPrecent[i])
    //                data[i, 25] = zc[i] * (1 - jfPrecent[i] / 100) * (1 - 自产焦水分 / 100) + ld[i] * (1 - jfPrecent[i] / 100) * (1 - 落地焦水分 / 100);  //入炉焦炭
    //            else
    //                data[i, 25] = zc[i] * (1 - 自产焦水分 / 100) + ld[i]  * (1 - 落地焦水分 / 100) - jfAmount[i];  //入炉焦炭

    //            data[6, 25] += data[i, 25];
    //        }
            

    //        //碎铁
    //        cmd.CommandText = "select gaolu,feitie from xiaohao where rq=:rq";
    //        dr = cmd.ExecuteReader();
    //        while (dr.Read())
    //        {
    //            if (!dr.IsDBNull(0))
    //            {
    //                int gaolu = dr.GetInt32(0);
    //                data[gaolu - 1, 23] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
    //            }

    //        }
    //        dr.Close();


    //        //工艺称总和
    //        data[5, 14]=data[0, 14] + data[1, 14] + data[2, 14] + data[3, 14] + data[4, 14];


    //        //计算全铁厂量,矿耗总和
    //        for (int i = 0; i < 6; i++)
    //        {
    //            data[i, 3] = data[i, 0] + data[i, 1]+data[i,2];
    //            data[i, 15] = data[i, 16] + data[i, 17] + data[i, 18] + data[i, 19] + data[i, 20] + data[i, 21];
    //            //if (data[5, 14] > 0)
    //            //{
    //                //data[i, 13] = data[5, 13] * data[i, 14] / data[5, 14];
    //                //data[i, 25] = data[5, 25] * data[i, 14] / data[5, 14];
    //            //}
    //            if (data[i, 3] > 0)
    //            {
    //                data[i, 26] = data[i, 25]*1000 / data[i, 3]; //七号称焦比
    //                data[i, 27] = (data[i, 25] + data[i, 22] * 0.8 + data[i, 24] * 0.9) *1000/ data[i, 3]; //综合焦比
    //                if (data[i, 13]>0)
    //                   data[i, 28] = data[i, 14]*(data[i, 25]/data[i, 13])*1000 / data[i, 3]; //工艺成焦比
    //               data[i, 33] = data[i, 22] * 1000 / data[i, 3];//煤比
    //               data[i, 32] = data[i, 23] * 1000 / data[i, 3];//碎铁比
    //            }
    //            if (data[i, 15] > 0)
    //            {
    //                data[i, 31] = data[i, 15] / data[i, 25]; //焦炭负荷
    //                data[i, 34] = (data[i, 3] -0.8*data[i, 23])*100/ data[i, 15]; //回收率
    //            }
    //            if (lf[i] > 0)
    //                data[i, 35] = 0.79 * fy[i]*100 / lf[i]/24; //富氧率
    //            fy[6] += fy[i];
    //            lf[6] += lf[i];

    //        }
    //        if (lf[6] > 0)
    //           data[6, 35] = 0.79 * fy[5]*100/lf[5]/24; //富氧率

    //        //计算全厂总和
    //        for (int i = 0; i < 26; i++)
    //        {
    //            data[6, i] = data[0, i] + data[1, i] + data[2, i] + data[3, i] + data[4, i] + data[5, i];
    //        }
    //        for (int i = 0; i < 7; i++)
    //        {
    //            data[i, 5] = data[i, 4] / rq.Day;
    //        }
    //        if (data[6, 3] > 0)
    //        {
    //            data[6, 26] = data[6, 25]*1000 / data[6, 3];
    //            data[6, 27] = (data[6, 25] + data[6, 22] * 0.8 + data[6, 24] * 0.9)*1000 / data[6, 3]; //综合焦比
    //            if (data[6, 13] > 0)
    //                data[6, 28] = data[6, 14] * (data[6, 25] / data[6, 13]) * 1000 / data[6, 3]; //工艺成焦比
    //            data[6, 33] = data[6, 22] * 1000 / data[6, 3];//煤比
    //            data[6, 32] = data[6, 23] * 1000 / data[6, 3];//碎铁比
    //        }
    //        if (data[6, 15] > 0)
    //        {
    //            data[6, 31] = data[6, 15] / data[6, 25]; //焦炭负荷
    //            data[6, 34] = (data[6, 3] - 0.8 * data[6, 23])*100 / data[6, 15]; //回收率
    //        }
    //        //实物系数
    //        data[0, 29] = data[0, 3] / Constants.volume1;
    //        data[1, 29] = data[1, 3] / Constants.volume2;
    //        data[2, 29] = data[2, 3] / Constants.volume3;
    //        data[3, 29] = data[3, 3] / Constants.volume4;
    //        data[4, 29] = data[4, 3] / Constants.volume5;
    //        data[5, 29] = data[5, 3] / Constants.volume6;
    //        data[6, 29] = data[6, 3] / Constants.volume;
    //        //冶炼强度
    //        data[0, 30] = data[0, 25] / Constants.volume1;
    //        data[1, 30] = data[1, 25] / Constants.volume2;
    //        data[2, 30] = data[2, 25] / Constants.volume3;
    //        data[3, 30] = data[3, 25] / Constants.volume4;
    //        data[4, 30] = data[4, 25] / Constants.volume5;
    //        data[5, 30] = data[5, 25] / Constants.volume5;
    //        data[6, 30] = data[6, 25] / Constants.volume;

    //        cn.Close();
    //    }
    //}
}
