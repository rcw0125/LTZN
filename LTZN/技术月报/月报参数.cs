using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;

namespace LTZN.技术月报
{
    class 月报参数
    {
        public int 本月天数 = 30;
        public double[] 高炉容积 = new double[6] { Constants.volume1, Constants.volume2, Constants.volume3, Constants.volume4, Constants.volume5, Constants.volume };

        public double[] 全铁产量 = new double[6];
        public double[] 合格铁 = new double[6];
        public double[] 制钢铁 = new double[6];
        public double[] 铸造铁 = new double[6];
        public double[] 一级铁 = new double[6];
        public double[] 送炼钢优质铁 = new double[6];
        public double[] 折算产量 = new double[6];
        
        public double[] 铁Fe = new double[6];
        public double[] 铁Si = new double[6];
        public double[] 铁Mn = new double[6];
        public double[] 铁P = new double[6];
        public double[] 铁S = new double[6];
        public double[] 渣CaO = new double[6];
        public double[] 渣SiO2 = new double[6];
        public double[] 渣Al2O3 = new double[6];
        public double[] 渣MgO = new double[6];
        public double[] 渣FeO = new double[6];
        public double[] 渣S = new double[6];

        public double[] 休风 = new double[6];
        public double[] 大中休 = new double[6];
        public double[] 慢风 = new double[6];

        public double[] 煤粉 = new double[6];
        public double[] 焦丁 = new double[6];
        public double[] 干毛焦 = new double[6];
        public double[] 入炉焦炭 = new double[6];
        public double[] 干焦粉= new double[6];
        public double[] 七号称焦炭 = new double[6];
        public double[] 工艺称焦炭 = new double[6];
        public double[] 工艺称焦比 = new double[6]; 


        public double[] 原料总耗 = new double[6];
        public double[] 机烧 = new double[6];
        public double[] 竖球 = new double[6];
        public double[] 其它熟料 = new double[6];
        public double[] 本溪矿 = new double[6];
        public double[] 其它生料 = new double[6];

        public double 机烧品位 = 0;                       
        public double 竖球品位 = 0;                       
        public double 其它熟料品位 = 0;                   
        public double 本溪矿品位 = 0;                    
        public double 其它生料品位 = 0;                  
        public double 机烧CaO = 0;                       
        public double 竖球CaO = 0;                       
        public double 其它熟料CaO = 0;                    
        public double 本溪矿CaO = 0;                     
        public double 其它生料CaO = 0;  
        public double 焦炭灰份 = 0;                    
        public double 煤粉灰份 = 0;
        public double 七号称水份 = 0;

        public double[] 废铁 = new double[6];
        public double[] 瓦斯灰 = new double[6];

        public double[] 冷风流量 = new double[6];
        public double[] 富氧流量 = new double[6];
        public double[] 平均风温 = new double[6];
        public double[] 热风压力 = new double[6];
        public double[] 炉顶温度 = new double[6];
        public double[] 炉顶压力 = new double[6];

        public double[] 悬料次数 = new double[6];
        public double[] 坐料次数 = new double[6];
        public double[] 崩料次数 = new double[6];

        public double[] 风口损坏大 = new double[6];
        public double[] 风口损坏中 = new double[6];
        public double[] 风口损坏小 = new double[6];
        public double[] 渣口损坏大 = new double[6];
        public double[] 渣口损坏中 = new double[6];
        public double[] 渣口损坏小 = new double[6];

        public double[] 制钢铁Si偏差 = new double[6];
        public double[] 铸造铁Si偏差 = new double[6];

        public double[] CO2 = new double[6];
        public double[] CO = new double[6];
        public double[] 深料线 = new double[6];
        public double[] 全部料线 = new double[6];
        public double[] 高压操作时间 = new double[6];         //单位 小时

        public double 调整机烧 = 0;
        public double 调整竖球 = 0;
        public double 调整其它熟料 = 0;
        public double 调整本溪矿 = 0;
        public double 调整其它生料 = 0;

        public double 调整干毛焦 = 0;
        public double 调整入炉焦炭 = 0;
        public double 调整焦丁 = 0;
        public double 调整煤粉 = 0;
        public double 调整富氧量 = 0;

        public double[] 高炉有效炉容 = new double[6];
        public double[] 合格率 = new double[6];
        public double[] 一级品率 = new double[6];
        public double[] 送炼钢优质铁水率 = new double[6];
        public double[] 高炉利用系数 = new double[6];
        public double[] 入炉焦炭冶炼强度 = new double[6];
        public double[] 综合焦炭冶炼强度 = new double[6];
        public double[] 折算综合焦比 = new double[6];
        public double[] 折合入炉焦比 = new double[6];
        public double[] 入炉综合品位 = new double[6];
        public double[] 熟料比 = new double[6];
        public double[] 入炉焦炭负荷 = new double[6];
        public double[] 综合焦炭负荷 = new double[6];
        public double[] 休风小时 = new double[6];
        public double[] 休风率 = new double[6];
        public double[] 慢风率 = new double[6];
        public double[] CO2率 = new double[6];
        public double[] 深料线率 = new double[6];
        public double[] 回收率 = new double[6];
        public double[] 富氧率 = new double[6];
        public double[] 高压率 = new double[6];
        public double[] 理论渣量 = new double[6];
        public double[] 渣铁比 = new double[6];
        public double[] 灰铁比 = new double[6];
        public double[] 原料单耗 = new double[6];
        public double[] 废铁单耗 = new double[6];
        public double[] 金属收得率 = new double[6];
        public double[] 综合焦炭 = new double[6];
        public double[] 综合焦炭单耗 = new double[6];
        public double[] 入炉焦炭单耗 = new double[6];
        public double[] 焦丁单耗 = new double[6];
        public double[] 煤粉单耗 = new double[6];
        public double[] 燃料比 = new double[6];
        public double[] 碱度 = new double[6];
        public double[] 收入含铁 = new double[6];
        public double[] 支出含铁 = new double[6];
        public double[] 有效工作时间 = new double[6];
        public double[] 高炉实际炉容 = new double[6];
        public double[] 入炉矿含铁 = new double[6];

        public void getDataBy(DateTime rq)
        {
            本月天数 = DateTime.DaysInMonth(rq.Year, rq.Month);

            OracleConnection cn = new OracleConnection();
            cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
            cn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = cn;
            DateTime temprq = new DateTime(rq.Year, rq.Month, 本月天数);

            cmd.CommandText = "select P01单位,P22平均风温,P23炉顶温度,P24热风压力,P25炉顶压力 from 技术日报 where P日期=to_date('" + temprq.ToString("yyyy-M-dd") + "','YYYY-MM-DD') and p02项目='累计'";
            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                string danwei = dr.IsDBNull(0) ? "" : dr.GetString(0);
                int i = 6;
                switch (danwei)
                {
                    case "1#":
                        i = 0;
                        break;
                    case "2#":
                        i = 1;
                        break;
                    case "3#":
                        i = 2;
                        break;
                    case "4#":
                        i = 3;
                        break;
                    case "5#":
                        i = 4;
                        break;
                    case "Q":
                        i = 5;
                        break;
                }
                if (i < 6)
                {
                    平均风温[i] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    炉顶温度[i] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                    热风压力[i] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                    炉顶压力[i] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                }
            }
            dr.Close();
            cmd.Parameters.Add(":rq", OracleType.DateTime).Value = rq;
           
           /*
            cmd.CommandText = "select gaolu,trunc(avg(fengwen),0),trunc(avg(ludingwendu),0),trunc(avg(refengyali),0),trunc(avg(ludingyali),0)  from xiaohao,全厂日消耗 where 全厂日消耗.P01日期=xiaohao.RQ and trunc(RQ,'MONTH')=trunc(:rq,'MONTH') and RQ<=:rq  group by gaolu";
            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0); if (gaolu > 5) continue;
                    平均风温[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    炉顶温度[gaolu - 1] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                    热风压力[gaolu - 1] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                    炉顶压力[gaolu - 1] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                }
            }
            dr.Close();
            for (int i = 0; i < 5; i++)
            {
                热风压力[5] += 热风压力[i];
                炉顶温度[5] += 炉顶温度[i];
                炉顶压力[5] += 炉顶压力[i];
            }
            平均风温[5] = (平均风温[0] * 350 + 平均风温[1] * 350 + 平均风温[2] * 350 + 平均风温[3] * 400 + 平均风温[4] * 420) / 1870;
            热风压力[5] = 热风压力[5] / 5;
            炉顶温度[5] = 炉顶温度[5] / 5;
            炉顶压力[5] = 炉顶压力[5] / 5;
*/
            //全铁产量
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where  trunc(zdsj,'MONTH')=trunc(:rq,'MONTH')  and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0); if (gaolu > 5) continue;
                    全铁产量[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    全铁产量[5] += 全铁产量[gaolu - 1];
                }
            }
            dr.Close();

            cmd.CommandText = "select 高炉,sum(工艺称) from 原料消耗 where trunc(日期,'MONTH')=trunc(:rq,'MONTH') group by 高炉";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                 if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0); if (gaolu > 5) continue;
                    工艺称焦炭[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    工艺称焦炭[5] += 工艺称焦炭[gaolu - 1];
                }
            }
            dr.Close();

            //合格铁
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where 合格铁(fesi,fes)=1 and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH')  and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0); if (gaolu > 5) continue;
                    合格铁[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    合格铁[5] += 合格铁[gaolu - 1];
                }
            }
            dr.Close();

            //制钢铁
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where fesi<=1.25 and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH')  and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0); if (gaolu > 5) continue;
                    制钢铁[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    制钢铁[5] += 制钢铁[gaolu - 1];
                }
            }
            dr.Close();

            //制钢铁SI标准偏差
            cmd.CommandText = "select gaolu,round(stddev(FESI),2) from ddluci where fesi<=1.25 and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH')  and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0); if (gaolu > 5) continue;
                    制钢铁Si偏差[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                }
            }
            dr.Close();

            //铸造铁SI标准偏差
            cmd.CommandText = "select gaolu,round(stddev(FESI),2) from ddluci where fesi>1.25 and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH')  and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0); if (gaolu > 5) continue;
                    铸造铁Si偏差[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                }
            }
            dr.Close();

            //全厂制钢铁SI标准偏差
            cmd.CommandText = "select round(stddev(FESI),2) from ddluci where fesi<=1.25 and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH')  and dksj is not null";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    制钢铁Si偏差[5] = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
                }
            }
            dr.Close();

            //全厂铸造铁SI标准偏差
            cmd.CommandText = "select round(stddev(FESI),2) from ddluci where fesi>1.25 and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH')  and dksj is not null";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    铸造铁Si偏差[5] = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
                }
            }
            dr.Close();

            //折算产量
            cmd.CommandText = "select gaolu,sum(折算产量(feliang,fesi,fes)) from ddluci where  trunc(zdsj,'MONTH')=trunc(:rq,'MONTH')  and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0); if (gaolu > 5) continue;
                    折算产量[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    折算产量[5] += 折算产量[gaolu - 1];
                }
            }
            dr.Close();

            //一级铁
            cmd.CommandText = "select gaolu,sum(FELIANG) from ddluci where ((FESI<=1.25 and FES<=0.03) or (FESI>1.25 and FESI<1.6 and FES<=0.04) or (FESI>=1.6 and FES<=0.03)) and  trunc(ZDSJ,'MONTH')=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0); if (gaolu > 5) continue;
                    一级铁[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    一级铁[5] += 一级铁[gaolu - 1];
                }
            }
            dr.Close();


            //送炼钢优质铁
            cmd.CommandText = "select gaolu,sum(FELIANG) from ddluci where FESI<=0.7 and FES<=0.04 and  trunc(ZDSJ,'MONTH')=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0); if (gaolu > 5) continue;
                    送炼钢优质铁[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    送炼钢优质铁[5] += 送炼钢优质铁[gaolu - 1];
                }
            }
            dr.Close();

            //铁成分
            cmd.CommandText = "select gaolu,trunc(avg(FESI),2),trunc(avg(FEMn),2),trunc(avg(FEP),3),trunc(avg(FES),3),round(avg(100-NVL(FEC,0)-NVL(FESI,0)-NVL(FEMN,0)-NVL(FEP,0)-NVL(FES,0)-NVL(FETI,0)),2)  from ddluci where  trunc(zdsj,'MONTH')=trunc(:rq,'MONTH')  and dksj is not null group by gaolu";
            cmd.Parameters.Add(":rq", OracleType.DateTime).Value = rq;
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0); if (gaolu > 5) continue;
                    铁Si[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    铁Mn[gaolu - 1] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                    铁P[gaolu - 1] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                    铁S[gaolu - 1] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                    铁Fe[gaolu - 1] = dr.IsDBNull(5) ? 0 : dr.GetDouble(5);
                }
            }
            dr.Close();

            //渣成分
            cmd.CommandText = "select gaolu,trunc(avg(ZHACAO),2),trunc(avg(ZHASIO2),2),trunc(avg(ZHAAL2O3),2),trunc(avg(ZHAMGO),2),trunc(avg(ZHAS),2)  from ddluci where trunc(zdsj,'MONTH')=trunc(:rq,'MONTH')  and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0); if (gaolu > 5) continue;
                    渣CaO[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    渣SiO2[gaolu - 1] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                    渣Al2O3[gaolu - 1] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                    渣MgO[gaolu - 1] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                    渣S[gaolu - 1] = dr.IsDBNull(5) ? 0 : dr.GetDouble(5);
                }
            }
            dr.Close();

            //铁成分
            cmd.CommandText = "select trunc(avg(FESI),2),trunc(avg(FEMn),2),trunc(avg(FEP),3),trunc(avg(FES),3),round(avg(100-NVL(FEC,0)-NVL(FESI,0)-NVL(FEMN,0)-NVL(FEP,0)-NVL(FES,0)-NVL(FETI,0)),2)  from ddluci where  trunc(zdsj,'MONTH')=trunc(:rq,'MONTH')  and dksj is not null";
            cmd.Parameters.Add(":rq", OracleType.DateTime).Value = rq;
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    铁Si[5] = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
                    铁Mn[5] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    铁P[5] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                    铁S[5] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                    铁Fe[5] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                }
            }
            dr.Close();

            //渣成分
            cmd.CommandText = "select trunc(avg(ZHACAO),2),trunc(avg(ZHASIO2),2),trunc(avg(ZHAAL2O3),2),trunc(avg(ZHAMGO),2),trunc(avg(ZHAS),2)  from ddluci where trunc(zdsj,'MONTH')=trunc(:rq,'MONTH')  and dksj is not null";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    渣CaO[5] = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
                    渣SiO2[5] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    渣Al2O3[5] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                    渣MgO[5] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                    渣S[5] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                }
            }
            dr.Close();



            cmd.CommandText = "select P01单位,sum(P38干毛焦),sum(P27入炉焦炭总耗),sum(P29煤粉总耗),sum(P31焦丁总耗),SUM(P12原料矿机烧),SUM(P13原料矿竖炉球),SUM(P15原料矿其它熟料),SUM(P16原料矿本溪矿),SUM(P17原料矿其它生料),SUM(P18废铁总耗)"
                              + ",ROUND(avg(p22平均风温),0),ROUND(avg(p23炉顶温度),0),ROUND(avg(p24热风压力),0),ROUND(avg(p25炉顶压力),0)"
                               + ",SUM(P53风口损坏数大),SUM(P54风口损坏数中),SUM(P55风口损坏数小),SUM(P56渣口损坏数大),SUM(P57渣口损坏数中),SUM(P58渣口损坏数小)"
                               + ",sum(P50坐料次数),sum(P51悬料次数),sum(P52崩料次数),sum(P48休风情况),sum(P49慢风) from 技术日报 where trunc(P日期,'MONTH')=trunc(:rq,'MONTH') and p02项目='本日' group by P01单位";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                string danwei = dr.IsDBNull(0) ? "" : dr.GetString(0);
                int i = 6;
                switch (danwei)
                {
                    case "1#":
                        i = 0;
                        break;
                    case "2#":
                        i = 1;
                        break;
                    case "3#":
                        i = 2;
                        break;
                    case "4#":
                        i = 3;
                        break;
                    case "5#":
                        i = 4;
                        break;
                    case "Q":
                        i = 5;
                        break;
                }
                if (i < 6)
                {
                    干毛焦[i] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    入炉焦炭[i] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                    煤粉[i] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                    焦丁[i] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                    机烧[i] = dr.IsDBNull(5) ? 0 : dr.GetDouble(5);
                    竖球[i] = dr.IsDBNull(6) ? 0 : dr.GetDouble(6);
                    其它熟料[i] = dr.IsDBNull(7) ? 0 : dr.GetDouble(7);
                    本溪矿[i] = dr.IsDBNull(8) ? 0 : dr.GetDouble(8);
                    其它生料[i] = dr.IsDBNull(9) ? 0 : dr.GetDouble(9);
                    废铁[i] = dr.IsDBNull(10) ? 0 : dr.GetDouble(10);
                    //平均风温[i] = dr.IsDBNull(11) ? 0 : dr.GetDouble(11);
                    //炉顶温度[i] = dr.IsDBNull(12) ? 0 : dr.GetDouble(12);
                    //热风压力[i] = dr.IsDBNull(13) ? 0 : dr.GetDouble(13);
                    //炉顶压力[i] = dr.IsDBNull(14) ? 0 : dr.GetDouble(14);
                    风口损坏大[i] = dr.IsDBNull(15) ? 0 : dr.GetDouble(15);
                    风口损坏中[i] = dr.IsDBNull(16) ? 0 : dr.GetDouble(16);
                    风口损坏小[i] = dr.IsDBNull(17) ? 0 : dr.GetDouble(17);
                    渣口损坏大[i] = dr.IsDBNull(18) ? 0 : dr.GetDouble(18);
                    渣口损坏中[i] = dr.IsDBNull(19) ? 0 : dr.GetDouble(19);
                    渣口损坏小[i] = dr.IsDBNull(20) ? 0 : dr.GetDouble(20);
                    坐料次数[i] = dr.IsDBNull(21) ? 0 : dr.GetDouble(21);
                    悬料次数[i] = dr.IsDBNull(22) ? 0 : dr.GetDouble(22);
                    崩料次数[i] = dr.IsDBNull(23) ? 0 : dr.GetDouble(23);
                    休风[i] = dr.IsDBNull(24) ? 0 : dr.GetDouble(24);
                    慢风[i] = dr.IsDBNull(25) ? 0 : dr.GetDouble(25);
                    七号称焦炭[i] = 干毛焦[i];
                }
            }
            dr.Close();

            //大中修
            cmd.CommandText = "select 高炉,sum(间隔) from 休风 where trunc(时间,'MONTH')=trunc(:rq,'MONTH') and (分类='大修' or 分类='中修')  and  间隔>0　and 间隔 is not null group by 高炉";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                int gaolu = dr.IsDBNull(0) ? 6 : dr.GetInt32(0);
                if (gaolu < 6) 大中休[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
            }
            大中休[5] = 大中休[0] + 大中休[1] + 大中休[2] + 大中休[3] + 大中休[4];
            dr.Close();

            cmd.CommandText = "select gaolu,round(avg(CO2),2),round(avg(CO),2),SUM(SHENLIAOPI),SUM(LIAOPI),SUM(FUYANG),SUM(LENGFENGLIULIANG*24),SUM(GAOYASHIJIAN) from XIAOHAO where trunc(RQ,'MONTH')=trunc(:rq,'MONTH') group by gaolu";
            dr = cmd.ExecuteReader();
            int CO2chushu = 0;
            int COchushu = 0;
            while (dr.Read())
            {
                int gaolu = dr.IsDBNull(0) ? 6 : dr.GetInt32(0);
                if (gaolu < 6)
                {
                    CO2[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    CO[gaolu - 1] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                    深料线[gaolu - 1] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                    全部料线[gaolu - 1] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                    富氧流量[gaolu - 1] = dr.IsDBNull(5) ? 0 : dr.GetDouble(5);
                    冷风流量[gaolu - 1] = dr.IsDBNull(6) ? 0 : dr.GetDouble(6);
                    高压操作时间[gaolu - 1] = dr.IsDBNull(7) ? 0 : dr.GetDouble(7);
                    CO2[5] += CO2[gaolu - 1];
                    if (CO2[gaolu - 1] > 0) CO2chushu++;
                    CO[5] += CO[gaolu - 1];
                    if (CO[gaolu - 1] > 0) COchushu++;
                    深料线[5] += 深料线[gaolu - 1];
                    全部料线[5] += 全部料线[gaolu - 1];
                    富氧流量[5] += 富氧流量[gaolu - 1];
                    冷风流量[5] += 冷风流量[gaolu - 1];
                    高压操作时间[5] += 高压操作时间[gaolu - 1];
                }
            }
            if (CO2chushu > 0) CO2[5] = CO2[5] / CO2chushu;
            if (COchushu > 0) CO[5] = CO[5] / COchushu;
            CO2[5] = double.Parse(CO2[5].ToString("N2"));
            CO[5] = double.Parse(CO[5].ToString("N2"));
            dr.Close();

            //七号称水份
            cmd.CommandText = "select ROUND(AVG((P02自产焦水分+P03落地焦水分)/2),2) from 全厂日消耗 where trunc(P01日期,'MONTH')=trunc(:rq,'MONTH')";
            dr = cmd.ExecuteReader();
            if (dr.Read())
            {

                七号称水份 = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
            }
            dr.Close();

            //调整参数
            cmd.Parameters.Clear();
            cmd.Parameters.Add(":nian", OracleType.Int32).Value = rq.Year;
            cmd.Parameters.Add(":yue", OracleType.Int32).Value = rq.Month;
            cmd.CommandText = "select 名称,重量,TFE,CAO from 每月矿 where 年=:nian and 月=:yue";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                string mc = dr.IsDBNull(0) ? "" : dr.GetString(0);
                switch (mc)
                {
                    case "机烧":
                        调整机烧 = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                        机烧品位 = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                        机烧CaO = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                        break;
                    case "竖炉球":
                        调整竖球 = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                        竖球品位 = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                        竖球CaO = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                        break;
                    case "其它熟料":
                        调整其它熟料 = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                        其它熟料品位 = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                        其它熟料CaO = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                        break;
                    case "本溪矿":
                        调整本溪矿 = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                        本溪矿品位 = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                        本溪矿CaO = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                        break;
                    case "其它生料":
                        调整其它生料 = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                        其它生料品位 = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                        其它生料CaO = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                        break;
                    case "瓦斯灰":
                        瓦斯灰[5] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                        break;
                }
            }
            dr.Close();
            cmd.CommandText = "select 干毛焦,round(入炉焦炭,0),焦丁,焦炭灰份 from 每月焦 where 年=:nian and 月=:yue";
            dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                调整干毛焦 = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
                调整入炉焦炭 = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                调整焦丁 = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                焦炭灰份 = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
            }
            dr.Close();

            cmd.CommandText = "select 重量*(1-水份/100),灰份 from 每月煤 where 年=:nian and 月=:yue";
            dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                调整煤粉 = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
                煤粉灰份 = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
            }
            dr.Close();

            cmd.CommandText = "select 富氧量 from 每月能耗 where 年=:nian and 月=:yue";
            dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                调整富氧量 = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
            }
            dr.Close();
            cn.Close();

            分配(机烧, 调整机烧, "N0");
            分配(竖球, 调整竖球, "N0");
            分配(其它熟料, 调整其它熟料, "N0");
            分配(本溪矿, 调整本溪矿, "N0");
            分配(其它生料, 调整其它生料, "N0");
            分配(煤粉, 调整煤粉, "N0");
            分配(焦丁, 调整焦丁, "N0");
            分配(干毛焦, 调整干毛焦, "N0");
            分配(入炉焦炭, 调整入炉焦炭, "N0");
            分配(七号称焦炭, 调整干毛焦 / (1 - 七号称水份 / 100), "N0");
            分配(富氧流量, 调整富氧量, "N0");
            分配成整数(全铁产量);
            分配成整数(合格铁);
            分配成整数(制钢铁);
            分配成整数(折算产量);
            分配成整数(一级铁);
            分配成整数(送炼钢优质铁);

            计算铸造铁();
            计算干焦粉();
            计算原料总耗();
            计算综合焦炭();
            计算高炉有效炉容();
            计算合格率();
            计算一级品率();
            计算休风小时();
            计算送炼钢优质铁水率();
            计算有效工作时间();
            计算高炉实际炉容();
            计算高炉利用系数();
            计算入炉焦炭冶炼强度();
            计算综合焦炭冶炼强度();
            计算折算综合焦比();
            计算折合入炉焦比();
            计算入炉矿含铁();
            计算入炉综合品位();
            计算熟料比();
            计算入炉焦炭负荷();
            计算综合焦炭负荷();
            计算休风率();
            计算慢风率();
            计算CO2率();
            计算深料线率();
            计算回收率();
            计算富氧率();
            计算高压率();
            计算理论渣量();
            计算渣铁比();
            计算灰铁比();
            计算原料单耗();
            计算废铁单耗();
            计算收入含铁();
            计算支出含铁();
            计算金属收得率();
            计算综合焦炭单耗();
            计算入炉焦炭单耗();
            计算焦丁单耗();
            计算煤粉单耗();
            计算燃料比();
            计算碱度();
            计算工艺称焦比();
        }

        private void 分配成整数(double[] dArray)
        {
            double temp = Convert.ToInt32(dArray[5]);
            if (temp <= 0) return;
            dArray[5] = temp;
            dArray[4] = Convert.ToInt32(temp * dArray[4] / (dArray[0] + dArray[1] + dArray[2] + dArray[3] + dArray[4]));
            temp = temp - dArray[4];
            if (temp <= 0)
            {
                dArray[0] = 0;
                dArray[1] = 0;
                dArray[2] = 0;
                dArray[3] = 0;
                return;
            }
            dArray[3] = Convert.ToInt32(temp * dArray[3] / (dArray[0] + dArray[1] + dArray[2] + dArray[3]));
            temp = temp - dArray[3];
            if (temp <= 0)
            {
                dArray[0] = 0;
                dArray[1] = 0;
                dArray[2] = 0;
                return;
            }
            dArray[2] = Convert.ToInt32(temp * dArray[2] / (dArray[0] + dArray[1] + dArray[2]));
            temp = temp - dArray[2];
            if (temp <= 0)
            {
                dArray[0] = 0;
                dArray[1] = 0;
                return;
            }
            dArray[1] = Convert.ToInt32(temp * dArray[1] / (dArray[0] + dArray[1]));
            dArray[0] = temp - dArray[1];
        }
        private void 分配(double[] dArray, double shu, string format)
        {
            if (shu <= 0) return;
            double temp = double.Parse(shu.ToString(format));
            dArray[5] = temp;
            if (temp <= 0)
            {
                dArray[0] = 0;
                dArray[1] = 0;
                dArray[2] = 0;
                dArray[3] = 0;
                dArray[4] = 0;
                return;
            }
            dArray[4] = temp * dArray[4] / (dArray[0] + dArray[1] + dArray[2] + dArray[3] + dArray[4]);
            dArray[4] = double.Parse(dArray[4].ToString(format));
            temp = temp - dArray[4];
            if (temp <= 0)
            {
                dArray[0] = 0;
                dArray[1] = 0;
                dArray[2] = 0;
                dArray[3] = 0;
                return;
            }
            dArray[3] = temp * dArray[3] / (dArray[0] + dArray[1] + dArray[2] + dArray[3]);
            dArray[3] = double.Parse(dArray[3].ToString(format));
            temp = temp - dArray[3];
            if (temp <= 0)
            {
                dArray[0] = 0;
                dArray[1] = 0;
                dArray[2] = 0;
                return;
            }
            dArray[2] = temp * dArray[2] / (dArray[0] + dArray[1] + dArray[2]);
            dArray[2] = double.Parse(dArray[2].ToString(format));
            temp = temp - dArray[2];
            if (temp <= 0)
            {
                dArray[0] = 0;
                dArray[1] = 0;
                return;
            }
            dArray[1] = temp * dArray[1] / (dArray[0] + dArray[1]);
            dArray[1] = double.Parse(dArray[1].ToString(format));
            dArray[0] = temp - dArray[1];
        }
        private void 计算干焦粉()
        {
            for (int i = 0; i < 6; i++)
            {
                干焦粉[i] = 干毛焦[i] - 入炉焦炭[i];
            }
        }
        private void 计算原料总耗()
        {
            for (int i = 0; i < 6; i++)
            {
                原料总耗[i] = 机烧[i] + 竖球[i] + 其它熟料[i] + 本溪矿[i] + 其它生料[i];
            }
        }
        private void 计算综合焦炭()
        {
            for (int i = 0; i < 6; i++)
            {
                综合焦炭[i] = 入炉焦炭[i] + 0.8 * 煤粉[i] + 0.9 * 焦丁[i];
            }
            分配成整数(综合焦炭);
        }
        private void 计算高炉有效炉容()
        {
            高炉有效炉容[0] = Constants.volume1 * (本月天数 - 休风[0] / 1440);
            高炉有效炉容[1] = Constants.volume2 * (本月天数 - 休风[1] / 1440);
            高炉有效炉容[2] = Constants.volume3 * (本月天数 - 休风[2] / 1440);
            高炉有效炉容[3] = Constants.volume4 * (本月天数 - 休风[3] / 1440);
            高炉有效炉容[4] = Constants.volume5 * (本月天数 - 休风[4] / 1440);
            高炉有效炉容[5] = 高炉有效炉容[0] + 高炉有效炉容[1] + 高炉有效炉容[2] + 高炉有效炉容[3] + 高炉有效炉容[4];
            分配成整数(高炉有效炉容);
  
        }
        private void 计算铸造铁()
        {
            for(int i=0;i<6;i++)
            {
                铸造铁[i] = 全铁产量[i] - 制钢铁[i];
            }
        }
        private void 计算合格率()
        {
            for (int i = 0; i < 6; i++)
            {
                if (全铁产量[i] > 0)
                    合格率[i] = 合格铁[i] * 100 / 全铁产量[i];
                合格率[i]=double.Parse(合格率[i].ToString("N2"));
            }
        }
        private void 计算一级品率()
        {
            for (int i = 0; i < 6; i++)
            {
                if (全铁产量[i] > 0)
                    一级品率[i] = 一级铁[i] * 100 / 全铁产量[i];
                一级品率[i]=double.Parse(一级品率[i].ToString("N2"));
            }
        }
        private void 计算送炼钢优质铁水率()
        {
            for (int i = 0; i < 6; i++)
            {
                if (全铁产量[i] > 0)
                    送炼钢优质铁水率[i] = 送炼钢优质铁[i] * 100 / 全铁产量[i];
                送炼钢优质铁水率[i] = double.Parse(送炼钢优质铁水率[i].ToString("N2"));
            }
        }
        private void 计算高炉实际炉容()
        {
            高炉实际炉容[0] = Constants.volume1 * (本月天数-大中休[0]/24);
            高炉实际炉容[1] = Constants.volume2 * (本月天数-大中休[1]/24);
            高炉实际炉容[2] = Constants.volume3 * (本月天数 - 大中休[2] / 24);
            高炉实际炉容[3] = Constants.volume4 * (本月天数 - 大中休[3] / 24);
            高炉实际炉容[4] = Constants.volume5 * (本月天数 - 大中休[4] / 24);
            高炉实际炉容[5] = 高炉实际炉容[0] + 高炉实际炉容[1] + 高炉实际炉容[2] + 高炉实际炉容[3] + 高炉实际炉容[4];
        }
        private void 计算高炉利用系数()
        {
            for (int i = 0; i < 6; i++)
            {
                if (高炉实际炉容[i] > 0)
                    高炉利用系数[i] = 折算产量[i] / 高炉实际炉容[i];
                高炉利用系数[i] = double.Parse(高炉利用系数[i].ToString("N3"));
            }
        }
        private void 计算入炉焦炭冶炼强度(){
            for (int i = 0; i < 6; i++)
            {
                if (高炉有效炉容[i] > 0)
                    入炉焦炭冶炼强度[i] = 入炉焦炭[i] / 高炉有效炉容[i];
                入炉焦炭冶炼强度[i] = double.Parse(入炉焦炭冶炼强度[i].ToString("N2"));
            }
        }
        private void 计算综合焦炭冶炼强度()
        {
            for (int i = 0; i < 6; i++)
            {
                if (高炉有效炉容[i] > 0)
                    综合焦炭冶炼强度[i] = 综合焦炭[i] / 高炉有效炉容[i];
                综合焦炭冶炼强度[i] = double.Parse(综合焦炭冶炼强度[i].ToString("N2"));
            }
        }
        private void 计算折算综合焦比(){
            for (int i = 0; i < 6; i++)
            {
                if (折算产量[i]>0)
                    折算综合焦比[i]=综合焦炭[i] * 1000 / 折算产量[i];
                折算综合焦比[i] = double.Parse(折算综合焦比[i].ToString("N0"));
            }
        }
        private void 计算折合入炉焦比()
        {
            for (int i = 0; i < 6; i++)
            {
                if (折算产量[i] > 0)
                    折合入炉焦比[i]=入炉焦炭[i] * 1000 / 折算产量[i];
                折合入炉焦比[i] = double.Parse(折合入炉焦比[i].ToString("N0"));
            }
        }
        private void 计算入炉矿含铁()
        {
            for (int i = 0; i < 6; i++)
            {
                入炉矿含铁[i] = 机烧[i] * 机烧品位 / 100 + 竖球[i] * 竖球品位 / 100 + 其它熟料[i] * 其它熟料品位 / 100 + 本溪矿[i] * 本溪矿品位 / 100 + 其它生料[i] * 其它生料品位 / 100;
            }
            分配成整数(入炉矿含铁);
        }
        private void 计算入炉综合品位()
        {
            for (int i = 0; i < 6; i++)
            {
                if (原料总耗[i] > 0)
                    入炉综合品位[i] = 入炉矿含铁[i]*100 / 原料总耗[i];
                入炉综合品位[i] = double.Parse(入炉综合品位[i].ToString("N2"));
            }
        }
        private void 计算熟料比(){
            for (int i = 0; i < 6; i++)
            {
                if (原料总耗[i] > 0)
                    熟料比[i] = (机烧[i] + 竖球[i] + 其它熟料[i]) * 100 / 原料总耗[i];
                熟料比[i] = double.Parse(熟料比[i].ToString("N2"));
            }
        }
        private void 计算入炉焦炭负荷()
        {
            for (int i = 0; i < 6; i++)
            {
                if (入炉焦炭[i] > 0)
                    入炉焦炭负荷[i] = 原料总耗[i] / 入炉焦炭[i];
                入炉焦炭负荷[i] = double.Parse(入炉焦炭负荷[i].ToString("N2"));
            }

        }
        private void 计算综合焦炭负荷()
        {
            for (int i = 0; i < 6; i++)
            {
                if (综合焦炭[i] > 0)
                    综合焦炭负荷[i] = 原料总耗[i] / 综合焦炭[i];
                综合焦炭负荷[i] = double.Parse(综合焦炭负荷[i].ToString("N2"));
            }
        }
        private void 计算休风小时()
        {
            for (int i = 0; i < 6; i++)
            {
                休风小时[i] = 休风[i] / 60;
                大中休[i] = 大中休[i] / 60;
                休风小时[i] = 休风小时[i] - 大中休[i];
            }
            分配成整数(休风小时);
            分配成整数(大中休);
        }
        private void 计算休风率()
        {
            for (int i = 0; i < 5; i++)
            {
                if (本月天数 * 24 - 大中休[i] > 0)
                {
                    休风率[i] = 休风小时[i] * 100 / (本月天数 * 24 - 大中休[i]);
                    休风率[i] = double.Parse(休风率[i].ToString("N2"));
                }
                else
                    休风率[i] = 0;
            }
            if (本月天数 * 5 * 24 - 大中休[5] > 0)
            {
                休风率[5] = 休风小时[5] * 100 / (本月天数 * 5 * 24 - 大中休[5]);
                休风率[5] = double.Parse(休风率[5].ToString("N2"));
            }
            else
                休风率[5] = 0;
        }
        private void 计算慢风率()
        {
            for (int i = 0; i < 5; i++)
            {
                if (本月天数 * 24 - 大中休[i] > 0)
                {
                    慢风率[i] = 慢风[i] * 100 / (本月天数 * 24 - 大中休[i]);
                    慢风率[i] = double.Parse(慢风率[i].ToString("N2"));
                }
                else
                    慢风率[i] = 0;
            }
            if (本月天数 * 5 * 24 - 大中休[5] > 0)
            {
                慢风率[5] = 慢风[5] * 100 / (本月天数 * 5 * 24 - 大中休[5]);
                慢风率[5] = double.Parse(慢风率[5].ToString("N2"));
            }
            else
                慢风率[5] = 0;
        }
        private void 计算CO2率(){
            for (int i = 0; i < 6; i++)
            {
                if (CO2[i] > 0)
                    CO2率[i] = CO2[i] * 100 / (CO2[i] + CO[i]);
                CO2率[i] = double.Parse(CO2率[i].ToString("N0"));
            }
        }
        private void 计算深料线率()
        {
            for (int i = 0; i < 6; i++)
            {
                    if (全部料线[i] > 0)
                        深料线率[i] = 深料线[i] * 100 / 全部料线[i];
                    深料线率[i] = double.Parse(深料线率[i].ToString("N2"));
            }
        }
        private void 计算回收率(){
            for (int i = 0; i < 6; i++)
            {
                if (原料总耗[i] > 0)
                    回收率[i] = (全铁产量[i] - 0.8 * 废铁[i]) * 100 / 原料总耗[i];
                回收率[i] = double.Parse(回收率[i].ToString("N2"));
            }
        }
        private void 计算富氧率()
        {
            for (int i = 0; i < 6; i++)
            {
                if (冷风流量[i] > 0)
                    富氧率[i] = ((冷风流量[i] * 0.21 + 富氧流量[i]) / (冷风流量[i] + 富氧流量[i]) - 0.21) * 100;
                  //  富氧率[i] = 0.79 * 富氧流量[i] * 100 / 冷风流量[i];
                富氧率[i] = double.Parse(富氧率[i].ToString("N2"));
            }
        }
        private void 计算有效工作时间()
        {
            for (int i = 0; i < 5; i++)
            {
                有效工作时间[i] = 本月天数 * 24-大中休[i];
            }
            有效工作时间[5] = 本月天数 * 5 * 24 -大中休[5];

        }
        private void 计算高压率()
        {
            for (int i = 0; i < 6; i++)
            {
                if (有效工作时间[i] > 0)
                    高压率[i] = 高压操作时间[i] * 100 / 有效工作时间[i];
                高压率[i] = double.Parse(高压率[i].ToString("N2"));
            }
        }
        private void 计算理论渣量()
        {
            for (int i = 0; i < 6; i++)
            {
                if (渣CaO[5] > 0)
                    理论渣量[i] = (机烧[i] * 机烧CaO + 竖球[i] * 竖球CaO + 其它熟料[i] * 其它熟料CaO + 本溪矿[i] * 本溪矿CaO + 其它生料[i] * 其它生料CaO + 入炉焦炭[i] * 焦炭灰份 * 0.0379 + 煤粉[i] * 煤粉灰份 * 0.0224) / 渣CaO[5];
            }
            分配成整数(理论渣量);
        }
        private void 计算渣铁比(){
            for (int i = 0; i < 6; i++)
            {
                if (全铁产量[i] > 0)
                    渣铁比[i] = 理论渣量[i] * 1000 / 全铁产量[i];
                渣铁比[i] = double.Parse(渣铁比[i].ToString("N0"));
            }
        }
        private void 计算灰铁比()
        {
            for (int i = 0; i < 5; i++)
            {
                瓦斯灰[i] = 高炉容积[i];
            }
            分配成整数(瓦斯灰);
            for (int i = 0; i < 6; i++)
            {
                if (全铁产量[i] > 0)
                    灰铁比[i] = 瓦斯灰[i]*1000 / 全铁产量[i];
                灰铁比[i] = double.Parse(灰铁比[i].ToString("N3"));
            }
        }
        private void 计算原料单耗()
        {
            for (int i = 0; i < 6; i++)
            {
                if (合格铁[i] > 0)
                    原料单耗[i] = 原料总耗[i] * 1000 / 合格铁[i];
                原料单耗[i] = double.Parse(原料单耗[i].ToString("N0"));
            }
        }
        private void 计算废铁单耗(){
            for (int i = 0; i < 6; i++)
            {
                if (合格铁[i] > 0)
                    废铁单耗[i] = 废铁[i] * 1000 / 合格铁[i];
                废铁单耗[i] = double.Parse(废铁单耗[i].ToString("N0"));
            }
        }
        private void 计算收入含铁()
        {
            for (int i = 0; i < 6; i++)
            {
                收入含铁[i] = 机烧[i] * 机烧品位 / 100 + 竖球[i] * 竖球品位 / 100 + 其它熟料[i] * 其它熟料品位 / 100 + 本溪矿[i] * 本溪矿品位 / 100 + 其它生料[i] * 其它生料品位 / 100 + 0.9 * 废铁[i];
            }
            分配成整数(收入含铁);
        }
        private void 计算支出含铁()
        {
            for (int i = 0; i < 6; i++)
            {
                支出含铁[i] = 全铁产量[i] * 铁Fe[i] / 100;
            }
            分配成整数(支出含铁);
        }
        private void 计算金属收得率(){
            for (int i = 0; i < 6; i++)
            {
                if (收入含铁[i] > 0)
                    金属收得率[i] = 支出含铁[i] * 100 / 收入含铁[i];
                金属收得率[i] = double.Parse(金属收得率[i].ToString("N2"));
            }

        }
        private void 计算综合焦炭单耗(){
            for (int i = 0; i < 6; i++)
            {
                if (合格铁[i] > 0)
                    综合焦炭单耗[i] = 综合焦炭[i] * 1000 / 合格铁[i];
                综合焦炭单耗[i] = double.Parse(综合焦炭单耗[i].ToString("N0"));
            }
        }
        private void 计算入炉焦炭单耗(){
            for (int i = 0; i < 6; i++)
            {
                if (合格铁[i] > 0)
                    入炉焦炭单耗[i] = 入炉焦炭[i] * 1000 / 合格铁[i];
                入炉焦炭单耗[i] = double.Parse(入炉焦炭单耗[i].ToString("N0"));
            }
        }
        private void 计算焦丁单耗(){
            for (int i = 0; i < 6; i++)
            {
                if (合格铁[i] > 0)
                    焦丁单耗[i] = 焦丁[i] * 1000 / 合格铁[i];
                焦丁单耗[i] = double.Parse(焦丁单耗[i].ToString("N0"));
            }
        }
        private void 计算煤粉单耗()
        {
            for (int i = 0; i < 6; i++)
            {
                if (合格铁[i] > 0)
                    煤粉单耗[i] = 煤粉[i] * 1000 / 合格铁[i];
                煤粉单耗[i] = double.Parse(煤粉单耗[i].ToString("N0"));
            }
        }
        private void 计算燃料比(){
            for (int i = 0; i < 6; i++)
            {
                if (合格铁[i] > 0)
                    燃料比[i] = (入炉焦炭[i] + 煤粉[i] + 焦丁[i]) * 1000 / 合格铁[i];
                燃料比[i] = double.Parse(燃料比[i].ToString("N0"));
            }
        }
        private void 计算碱度()
        {
            for (int i = 0; i < 6; i++)
            {
                if (渣SiO2[i] > 0)
                    碱度[i] = 渣CaO[i] / 渣SiO2[i];
                碱度[i] = double.Parse(碱度[i].ToString("N2"));
            }
        }
        private void 计算工艺称焦比()
        {
            for (int i = 0; i < 6; i++)
            {
                if (全铁产量[i] > 0)
                    工艺称焦比[i] = 工艺称焦炭[i] * 干毛焦[i]*1000 / 七号称焦炭[i] / 全铁产量[i];
                工艺称焦比[i] = double.Parse(工艺称焦比[i].ToString("N0"));
            }
        }

        public void 计算()
        {
            计算铸造铁();
            计算干焦粉();
            计算原料总耗();
            计算综合焦炭();
            计算高炉有效炉容();
            计算合格率();
            计算一级品率();
            计算送炼钢优质铁水率();
            计算高炉利用系数();
            计算入炉焦炭冶炼强度();
            计算综合焦炭冶炼强度();
            计算折算综合焦比();
            计算折合入炉焦比();
            计算入炉综合品位();
            计算熟料比();
            计算入炉焦炭负荷();
            计算综合焦炭负荷();
           // 计算休风小时();
            计算休风率();
            计算慢风率();
            计算CO2率();
            计算深料线率();
            计算回收率();
            计算富氧率();
            计算高压率();
            计算理论渣量();
            计算渣铁比();
            计算灰铁比();
            计算原料单耗();
            计算废铁单耗();
            计算金属收得率();
            计算综合焦炭单耗();
            计算入炉焦炭单耗();
            计算焦丁单耗();
            计算煤粉单耗();
            计算燃料比();
            计算碱度();
            计算工艺称焦比();
        }

    }
}
