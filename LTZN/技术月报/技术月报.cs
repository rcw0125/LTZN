using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;
using LTZN.CalFramework;
using LTZN.Repository;

namespace LTZN.技术月报
{
    public class 技术月报 : QueryModel
    {
        public 技术月报()
            : base("技术月报")
        {
           
        }

        public void getDataBy(int nian, int yue)
        {
            conn.Open();

            int 本月天数 = DateTime.DaysInMonth(nian, yue);
           
          

            int 上月累计天数 = 0; 
           

            DateTime curYear=new DateTime(nian,1,1);
            DateTime curMonth=new DateTime(nian,yue,1);
            DateTime lastDay = new DateTime(nian, yue, 本月天数);
            if (yue > 1)
            {
                上月累计天数 = curMonth.AddDays(-1).DayOfYear;
            }

            this.SetValue("本月.天数", 本月天数);
            this.SetValue("上月累计.天数", 上月累计天数);

            //2016-11-24 添加   P76萤石,P77球团矿,P78国内球团矿,P79进口球团矿,P80其它块矿,P81高钛球团矿,P82高品位钛球,P83其它熔剂
            StringBuilder shuliangSql = new StringBuilder();
            shuliangSql.Append("select P01单位,P22平均风温,P23炉顶温度,P24热风压力,P25炉顶压力");
            shuliangSql.Append(",P38干毛焦,P38干毛焦-P27入炉焦炭总耗,P29煤粉总耗,P31焦丁总耗,P12原料矿机烧,P13原料矿竖炉球,P15原料矿其它熟料,P16原料矿本溪矿,P17原料矿其它生料,P18废铁总耗");
            shuliangSql.Append(",P53风口损坏数大,P54风口损坏数中,P55风口损坏数小,P56渣口损坏数大,P57渣口损坏数中,P58渣口损坏数小");
            shuliangSql.Append(",P50坐料次数,P51悬料次数,P52崩料次数,P48休风情况,P49慢风");
            shuliangSql.Append(",P69PB块,P70纽曼块,P71钛球,P72锰矿,P73硅石,P74白云石,P75蛇纹石,P76萤石,P77球团矿,P78国内球团矿,P79进口球团矿,");
            shuliangSql.Append("P80其它块矿,P81高钛球团矿,P82高品位钛球,P83其它熔剂  from 技术日报 where P日期=:RQ and p02项目='累计'");
            
            OracleCommand shuliangCmd=new OracleCommand(shuliangSql.ToString(),conn);
            shuliangCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = lastDay;
            OracleDataReader shuliangDr = shuliangCmd.ExecuteReader();

            while (shuliangDr.Read())
            {
                string gaolustr = shuliangDr.GetString(0);
                if (gaolustr.Contains("#") && gaolustr.Length > 1)
                {
                    int gaolu;
                    if (int.TryParse(gaolustr.Substring(0, gaolustr.Length - 1), out gaolu))
                    {
                        //this.SetValue(string.Format("本月.高炉{0}.平均风温", gaolu), shuliangDr.IsDBNull(1) ? 0 : shuliangDr.GetDouble(1));
                        //this.SetValue(string.Format("本月.高炉{0}.炉顶温度", gaolu), shuliangDr.IsDBNull(2) ? 0 : shuliangDr.GetDouble(2));
                        //this.SetValue(string.Format("本月.高炉{0}.热风压力", gaolu), shuliangDr.IsDBNull(3) ? 0 : shuliangDr.GetDouble(3));
                        //this.SetValue(string.Format("本月.高炉{0}.炉顶压力", gaolu), shuliangDr.IsDBNull(4) ? 0 : shuliangDr.GetDouble(4));
                        this.SetValue(string.Format("本月.高炉{0}.统计干毛焦", gaolu), shuliangDr.IsDBNull(5) ? 0 : shuliangDr.GetDouble(5));
                        this.SetValue(string.Format("本月.高炉{0}.统计干焦粉", gaolu), shuliangDr.IsDBNull(6) ? 0 : shuliangDr.GetDouble(6));
                        this.SetValue(string.Format("本月.高炉{0}.统计煤粉", gaolu), shuliangDr.IsDBNull(7) ? 0 : shuliangDr.GetDouble(7));
                        this.SetValue(string.Format("本月.高炉{0}.统计焦丁", gaolu), shuliangDr.IsDBNull(8) ? 0 : shuliangDr.GetDouble(8));
                        this.SetValue(string.Format("本月.高炉{0}.统计机烧", gaolu), shuliangDr.IsDBNull(9) ? 0 : shuliangDr.GetDouble(9));
                        this.SetValue(string.Format("本月.高炉{0}.统计竖球", gaolu), shuliangDr.IsDBNull(10) ? 0 : shuliangDr.GetDouble(10));
                        this.SetValue(string.Format("本月.高炉{0}.统计其它熟料", gaolu), shuliangDr.IsDBNull(11) ? 0 : shuliangDr.GetDouble(11));
                        this.SetValue(string.Format("本月.高炉{0}.统计本溪矿", gaolu), shuliangDr.IsDBNull(12) ? 0 : shuliangDr.GetDouble(12));
                        this.SetValue(string.Format("本月.高炉{0}.统计其它生料", gaolu), shuliangDr.IsDBNull(13) ? 0 : shuliangDr.GetDouble(13));
                      //  this.SetValue(string.Format("本月.高炉{0}.废铁", gaolu), shuliangDr.IsDBNull(14) ? 0 : shuliangDr.GetDouble(14));
                        //this.SetValue(string.Format("本月.高炉{0}.风口损坏数大", gaolu), shuliangDr.IsDBNull(15) ? 0 : shuliangDr.GetDouble(15));
                        //this.SetValue(string.Format("本月.高炉{0}.风口损坏数中", gaolu), shuliangDr.IsDBNull(16) ? 0 : shuliangDr.GetDouble(16));
                        //this.SetValue(string.Format("本月.高炉{0}.风口损坏数小", gaolu), shuliangDr.IsDBNull(17) ? 0 : shuliangDr.GetDouble(17));
                        //this.SetValue(string.Format("本月.高炉{0}.渣口损坏数大", gaolu), shuliangDr.IsDBNull(18) ? 0 : shuliangDr.GetDouble(18));
                        //this.SetValue(string.Format("本月.高炉{0}.渣口损坏数中", gaolu), shuliangDr.IsDBNull(19) ? 0 : shuliangDr.GetDouble(19));
                        //this.SetValue(string.Format("本月.高炉{0}.渣口损坏数小", gaolu), shuliangDr.IsDBNull(20) ? 0 : shuliangDr.GetDouble(20));
                        //this.SetValue(string.Format("本月.高炉{0}.坐料次数", gaolu), shuliangDr.IsDBNull(21) ? 0 : shuliangDr.GetDouble(21));
                        //this.SetValue(string.Format("本月.高炉{0}.悬料次数", gaolu), shuliangDr.IsDBNull(22) ? 0 : shuliangDr.GetDouble(22));
                        //this.SetValue(string.Format("本月.高炉{0}.崩料次数", gaolu), shuliangDr.IsDBNull(23) ? 0 : shuliangDr.GetDouble(23));
                        //this.SetValue(string.Format("本月.高炉{0}.休风时间", gaolu), shuliangDr.IsDBNull(24) ? 0 : shuliangDr.GetDouble(24));
                       // P76萤石,P77球团矿,P78国内球团矿,P79进口球团矿,P80其它块矿,P81高钛球团矿,P82高品位钛球,P83其它熔剂
                        this.SetValue(string.Format("本月.高炉{0}.慢风时间", gaolu), shuliangDr.IsDBNull(25) ? 0 : shuliangDr.GetDouble(25));
                        this.SetValue(string.Format("本月.高炉{0}.统计PB块", gaolu), shuliangDr.IsDBNull(26) ? 0 : shuliangDr.GetDouble(26));
                        this.SetValue(string.Format("本月.高炉{0}.统计纽曼块", gaolu), shuliangDr.IsDBNull(27) ? 0 : shuliangDr.GetDouble(27));
                        this.SetValue(string.Format("本月.高炉{0}.统计钛球", gaolu), shuliangDr.IsDBNull(28) ? 0 : shuliangDr.GetDouble(28));
                        this.SetValue(string.Format("本月.高炉{0}.统计锰矿", gaolu), shuliangDr.IsDBNull(29) ? 0 : shuliangDr.GetDouble(29));
                        this.SetValue(string.Format("本月.高炉{0}.统计硅石", gaolu), shuliangDr.IsDBNull(30) ? 0 : shuliangDr.GetDouble(30));
                        this.SetValue(string.Format("本月.高炉{0}.统计白云石", gaolu), shuliangDr.IsDBNull(31) ? 0 : shuliangDr.GetDouble(31));
                        this.SetValue(string.Format("本月.高炉{0}.统计蛇纹石", gaolu), shuliangDr.IsDBNull(32) ? 0 : shuliangDr.GetDouble(32));

                        this.SetValue(string.Format("本月.高炉{0}.统计萤石", gaolu), shuliangDr.IsDBNull(33) ? 0 : shuliangDr.GetDouble(33));
                        this.SetValue(string.Format("本月.高炉{0}.统计球团矿", gaolu), shuliangDr.IsDBNull(34) ? 0 : shuliangDr.GetDouble(34));
                        this.SetValue(string.Format("本月.高炉{0}.统计国内球团矿", gaolu), shuliangDr.IsDBNull(35) ? 0 : shuliangDr.GetDouble(35));
                        this.SetValue(string.Format("本月.高炉{0}.统计进口球团矿", gaolu), shuliangDr.IsDBNull(36) ? 0 : shuliangDr.GetDouble(36));
                        this.SetValue(string.Format("本月.高炉{0}.统计其它块矿", gaolu), shuliangDr.IsDBNull(37) ? 0 : shuliangDr.GetDouble(37));
                        this.SetValue(string.Format("本月.高炉{0}.统计高钛球团矿", gaolu), shuliangDr.IsDBNull(38) ? 0 : shuliangDr.GetDouble(38));
                        this.SetValue(string.Format("本月.高炉{0}.统计高品位钛球", gaolu), shuliangDr.IsDBNull(39) ? 0 : shuliangDr.GetDouble(39));
                        this.SetValue(string.Format("本月.高炉{0}.统计其它熔剂", gaolu), shuliangDr.IsDBNull(40) ? 0 : shuliangDr.GetDouble(40));
                       
                    }
                }
                else if (gaolustr == "Q")
                {
                    //this.SetValue("本月.平均风温", shuliangDr.IsDBNull(1) ? 0 : shuliangDr.GetDouble(1));
                    //this.SetValue("本月.炉顶温度", shuliangDr.IsDBNull(2) ? 0 : shuliangDr.GetDouble(2));
                    //this.SetValue("本月.热风压力", shuliangDr.IsDBNull(3) ? 0 : shuliangDr.GetDouble(3));
                    //this.SetValue("本月.炉顶压力", shuliangDr.IsDBNull(4) ? 0 : shuliangDr.GetDouble(4));
                    this.SetValue("本月.统计干毛焦", shuliangDr.IsDBNull(5) ? 0 : shuliangDr.GetDouble(5));
                    this.SetValue("本月.统计干焦粉", shuliangDr.IsDBNull(6) ? 0 : shuliangDr.GetDouble(6));
                    this.SetValue("本月.煤粉", shuliangDr.IsDBNull(7) ? 0 : shuliangDr.GetDouble(7));
                    this.SetValue("本月.焦丁", shuliangDr.IsDBNull(8) ? 0 : shuliangDr.GetDouble(8));
                    this.SetValue("本月.机烧", shuliangDr.IsDBNull(9) ? 0 : shuliangDr.GetDouble(9));
                    this.SetValue("本月.竖球", shuliangDr.IsDBNull(10) ? 0 : shuliangDr.GetDouble(10));
                    this.SetValue("本月.其它熟料", shuliangDr.IsDBNull(11) ? 0 : shuliangDr.GetDouble(11));
                    this.SetValue("本月.本溪矿", shuliangDr.IsDBNull(12) ? 0 : shuliangDr.GetDouble(12));
                    this.SetValue("本月.其它生料", shuliangDr.IsDBNull(13) ? 0 : shuliangDr.GetDouble(13));
                   // this.SetValue("本月.废铁", shuliangDr.IsDBNull(14) ? 0 : shuliangDr.GetDouble(14));
                    //this.SetValue("本月.风口损坏数大", shuliangDr.IsDBNull(15) ? 0 : shuliangDr.GetDouble(15));
                    //this.SetValue("本月.风口损坏数中", shuliangDr.IsDBNull(16) ? 0 : shuliangDr.GetDouble(16));
                    //this.SetValue("本月.风口损坏数小",  shuliangDr.IsDBNull(17) ? 0 : shuliangDr.GetDouble(17));
                    //this.SetValue("本月.渣口损坏数大",  shuliangDr.IsDBNull(18) ? 0 : shuliangDr.GetDouble(18));
                    //this.SetValue("本月.渣口损坏数中", shuliangDr.IsDBNull(19) ? 0 : shuliangDr.GetDouble(19));
                    //this.SetValue("本月.渣口损坏数小",shuliangDr.IsDBNull(20) ? 0 : shuliangDr.GetDouble(20));
                    //this.SetValue("本月.坐料次数", shuliangDr.IsDBNull(21) ? 0 : shuliangDr.GetDouble(21));
                    //this.SetValue("本月.悬料次数", shuliangDr.IsDBNull(22) ? 0 : shuliangDr.GetDouble(22));
                    //this.SetValue("本月.崩料次数", shuliangDr.IsDBNull(23) ? 0 : shuliangDr.GetDouble(23));
                    //this.SetValue("本月.休风时间", shuliangDr.IsDBNull(24) ? 0 : shuliangDr.GetDouble(24));
                    this.SetValue("本月.慢风", shuliangDr.IsDBNull(25) ? 0 : shuliangDr.GetDouble(25));
                 //   this.SetValue("本月.钛球", shuliangDr.IsDBNull(28) ? 0 : shuliangDr.GetDouble(28));
                 //   this.SetValue("本月.统计钛球", shuliangDr.IsDBNull(28) ? 0 : shuliangDr.GetDouble(28));
                    this.SetValue("本月.萤石", shuliangDr.IsDBNull(33) ? 0 : shuliangDr.GetDouble(33));
                    this.SetValue("本月.球团矿", shuliangDr.IsDBNull(34) ? 0 : shuliangDr.GetDouble(34));
                    this.SetValue("本月.国内球团矿", shuliangDr.IsDBNull(35) ? 0 : shuliangDr.GetDouble(35));
                    this.SetValue("本月.进口球团矿", shuliangDr.IsDBNull(36) ? 0 : shuliangDr.GetDouble(36));
                    this.SetValue("本月.其它块矿", shuliangDr.IsDBNull(37) ? 0 : shuliangDr.GetDouble(37));
                    this.SetValue("本月.高钛球团矿", shuliangDr.IsDBNull(38) ? 0 : shuliangDr.GetDouble(38));
                    this.SetValue("本月.高品位钛球", shuliangDr.IsDBNull(39) ? 0 : shuliangDr.GetDouble(39));
                    this.SetValue("本月.其它熔剂", shuliangDr.IsDBNull(40) ? 0 : shuliangDr.GetDouble(40));

                }
            }
            shuliangDr.Close();

            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;
            cmd.Parameters.Add(":rq", OracleType.DateTime).Value = curMonth;
            
            //全铁产量
            cmd.CommandText = "select nvl(gaolu,9999),sum(feliang) from ddluci where  trunc(zdsj,'MONTH')=trunc(:rq,'MONTH')  and dksj is not null group by rollup(gaolu)";
            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    if (gaolu != 9999)
                        this.SetValue(string.Format("本月.高炉{0}.统计全铁产量", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                    else
                        this.SetValue("本月.统计全铁产量", dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();

            //制钢铁、铸造铁SI标准偏差
            cmd.CommandText = "select nvl(gaolu,9999),case when fesi<=1.25 then '制钢铁' else '铸造铁' end,sum(feliang),round(stddev(FESI),2) from ddluci where trunc(zdsj,'MONTH')=trunc(:rq,'MONTH')  and dksj is not null group by rollup(gaolu),case when fesi<=1.25 then '制钢铁' else '铸造铁' end";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    if (gaolu != 9999)
                    {
                        this.SetValue(string.Format("本月.高炉{0}.统计{1}", gaolu, dr.GetString(1)), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                        this.SetValue(string.Format("本月.高炉{0}.{1}{2}", gaolu, dr.GetString(1), "Si偏差"), dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                    }
                    else
                    {
                        this.SetValue(string.Format("本月.统计{0}", dr.GetString(1)), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                        this.SetValue(string.Format("本月.{0}{1}", dr.GetString(1), "Si偏差"), dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                    }
                }
            }
            dr.Close();

            //合格铁
            cmd.CommandText = "select nvl(gaolu,9999),case when 合格铁(fesi,fes)=1 then '合格铁' else '号外铁' end,sum(feliang) from ddluci where trunc(zdsj,'MONTH')=trunc(:rq,'MONTH')  and dksj is not null group by rollup(gaolu),case when 合格铁(fesi,fes)=1 then '合格铁' else '号外铁' end";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    if (gaolu != 9999)
                    {
                        this.SetValue(string.Format("本月.高炉{0}.统计{1}", gaolu, dr.GetString(1)), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                    }
                    else
                    {
                        this.SetValue(string.Format("本月.统计{0}", dr.GetString(1)), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                    }
                }
            }
            dr.Close();

            //一级铁
            cmd.CommandText = "select nvl(gaolu,9999),case when ((FESI<=1.25 and FES<=0.03) or (FESI>1.25 and FESI<1.6 and FES<=0.04) or (FESI>=1.6 and FES<=0.03)) then '一级铁' else '非一级铁' end,sum(FELIANG) from ddluci where   trunc(ZDSJ,'MONTH')=:rq and dksj is not null group by rollup(gaolu),case when ((FESI<=1.25 and FES<=0.03) or (FESI>1.25 and FESI<1.6 and FES<=0.04) or (FESI>=1.6 and FES<=0.03)) then '一级铁' else '非一级铁' end";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    if (gaolu != 9999)
                        this.SetValue(string.Format("本月.高炉{0}.统计{1}", gaolu, dr.GetString(1)), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                    else
                        this.SetValue(string.Format("本月.统计{0}", dr.GetString(1)), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                }
            }
            dr.Close();


            //送炼钢优质铁
            if(nian>=2013)
                cmd.CommandText = "select nvl(gaolu,9999),case when FESI<=0.7 and FES<=0.04 then '优质铁' else '非优质铁' end,sum(FELIANG) from ddluci where  trunc(ZDSJ,'MONTH')=:rq and dksj is not null group by rollup(gaolu),case when FESI<=0.7 and FES<=0.04 then '优质铁' else '非优质铁' end";
            else
            cmd.CommandText = "select nvl(gaolu,9999),case when FESI<=0.7 and FES<=0.04 then '优质铁' else '非优质铁' end,sum(FELIANG) from ddluci where  trunc(ZDSJ,'MONTH')=:rq and dksj is not null group by rollup(gaolu),case when FESI<=0.7 and FES<=0.04 then '优质铁' else '非优质铁' end";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                int gaolu = dr.GetInt32(0);
                if (gaolu != 9999)
                {
                    this.SetValue(string.Format("本月.高炉{0}.统计{1}", gaolu, dr.GetString(1)), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                }
                else
                {
                    this.SetValue(string.Format("本月.统计{0}", dr.GetString(1)), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                }
            }
            dr.Close();

            //折算产量
            cmd.CommandText = "select nvl(gaolu,9999),sum(折算产量(feliang,fesi,fes)) from ddluci where  trunc(zdsj,'MONTH')=trunc(:rq,'MONTH')  and dksj is not null group by rollup(gaolu)";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                int gaolu = dr.GetInt32(0);
                if (gaolu != 9999)
                {
                    this.SetValue(string.Format("本月.高炉{0}.统计折算产量", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
                else
                {
                    this.SetValue("本月.统计折算产量", dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();

            cmd.CommandText = "select 高炉,sum(工艺称) from 原料消耗 where trunc(日期,'MONTH')=trunc(:rq,'MONTH') group by 高炉";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("本月.高炉{0}.工艺称", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();

            //大中修，休风时间
            cmd.CommandText = "select nvl(高炉,9999),case when 分类='大修' or 分类='中修' then '大中修' else '休风时间' end,round(sum(间隔)/60,3) from 休风 where trunc(时间,'MONTH')=trunc(:rq,'MONTH') and  间隔>0　and 间隔 is not null group by rollup(高炉),case when 分类='大修' or 分类='中修' then '大中修' else '休风时间' end";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    if (gaolu != 9999)
                        this.SetValue(string.Format("本月.高炉{0}.{1}", gaolu, dr.GetString(1)), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                    else
                        this.SetValue(string.Format("本月.{0}", dr.GetString(1)), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                }
            }
            dr.Close();



            //七号称水份
            cmd.CommandText = "select ROUND(AVG((P02自产焦水分+P03落地焦水分)/2),2) from 全厂日消耗 where trunc(P01日期,'MONTH')=trunc(:rq,'MONTH')";
            dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                this.SetValue("本月.七号称水份", dr.IsDBNull(0) ? 0 : dr.GetDouble(0));
            }
            dr.Close();

            

            //本月消耗
            cmd.CommandText = "select nvl(gaolu,9999),round(avg(CO2),2),round(avg(CO),2),SUM(SHENLIAOPI),SUM(LIAOPI),SUM(FUYANG),SUM(LENGFENGLIULIANG*24),SUM(GAOYASHIJIAN),trunc(avg(fengwen),0),trunc(avg(ludingwendu),0),trunc(avg(refengyali),0),trunc(avg(ludingyali),0),sum(fengkoudatao),sum(fengkouzhongtao),sum(fengkouxiaotao),sum(zhakoudatao),sum(zhakouzhongtao),sum(zhakouxiaotao),sum(zuoliao),sum(xuanliao),sum(bengliao),sum(feitie) from XIAOHAO where trunc(RQ,'MONTH')=trunc(:rq,'MONTH') group by rollup(gaolu)";
            dr = cmd.ExecuteReader();
            int CO2chushu = 0;
            int COchushu = 0;
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu =  dr.GetInt32(0);
                    if (gaolu != 9999)
                    {
                        double CO2 = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                        double CO = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                        if (CO2 > 0) CO2chushu++;
                        if (CO > 0) COchushu++;

                        this.SetValue(string.Format("本月.高炉{0}.CO2", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                        this.SetValue(string.Format("本月.高炉{0}.CO", gaolu), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                        this.SetValue(string.Format("本月.高炉{0}.深料线", gaolu), dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                        this.SetValue(string.Format("本月.高炉{0}.全部料线", gaolu), dr.IsDBNull(4) ? 0 : dr.GetDouble(4));
                        this.SetValue(string.Format("本月.高炉{0}.统计富氧量", gaolu), dr.IsDBNull(5) ? 0 : dr.GetDouble(5));
                        this.SetValue(string.Format("本月.高炉{0}.冷风流量", gaolu), dr.IsDBNull(6) ? 0 : dr.GetDouble(6));
                        this.SetValue(string.Format("本月.高炉{0}.高压操作时间", gaolu), dr.IsDBNull(7) ? 0 : dr.GetDouble(7));
                        this.SetValue(string.Format("本月.高炉{0}.平均风温", gaolu), dr.IsDBNull(8) ? 0 : dr.GetDouble(8));
                        this.SetValue(string.Format("本月.高炉{0}.炉顶温度", gaolu), dr.IsDBNull(9) ? 0 : dr.GetDouble(9));
                        this.SetValue(string.Format("本月.高炉{0}.热风压力", gaolu), dr.IsDBNull(10) ? 0 : dr.GetDouble(10));
                        this.SetValue(string.Format("本月.高炉{0}.炉顶压力", gaolu), dr.IsDBNull(11) ? 0 : dr.GetDouble(11));
                        this.SetValue(string.Format("本月.高炉{0}.风口损坏数大", gaolu), dr.IsDBNull(12) ? 0 : dr.GetDouble(12));
                        this.SetValue(string.Format("本月.高炉{0}.风口损坏数中", gaolu), dr.IsDBNull(13) ? 0 : dr.GetDouble(13));
                        this.SetValue(string.Format("本月.高炉{0}.风口损坏数小", gaolu), dr.IsDBNull(14) ? 0 : dr.GetDouble(14));
                        this.SetValue(string.Format("本月.高炉{0}.渣口损坏数大", gaolu), dr.IsDBNull(15) ? 0 : dr.GetDouble(15));
                        this.SetValue(string.Format("本月.高炉{0}.渣口损坏数中", gaolu), dr.IsDBNull(16) ? 0 : dr.GetDouble(16));
                        this.SetValue(string.Format("本月.高炉{0}.渣口损坏数小", gaolu), dr.IsDBNull(17) ? 0 : dr.GetDouble(17));
                        this.SetValue(string.Format("本月.高炉{0}.坐料次数", gaolu), dr.IsDBNull(18) ? 0 : dr.GetDouble(18));
                        this.SetValue(string.Format("本月.高炉{0}.悬料次数", gaolu), dr.IsDBNull(19) ? 0 : dr.GetDouble(19));
                        this.SetValue(string.Format("本月.高炉{0}.崩料次数", gaolu), dr.IsDBNull(20) ? 0 : dr.GetDouble(20));
                      //  this.SetValue(string.Format("本月.高炉{0}.干焦粉", gaolu), dr.IsDBNull(21) ? 0 : dr.GetDouble(21));
                        this.SetValue(string.Format("本月.高炉{0}.废铁", gaolu), dr.IsDBNull(21) ? 0 : dr.GetDouble(21));

                    }
                    else
                    {
                        this.SetValue("本月.CO2", dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                        this.SetValue("本月.CO", dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                        this.SetValue("本月.深料线", dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                        this.SetValue("本月.全部料线", dr.IsDBNull(4) ? 0 : dr.GetDouble(4));
                        this.SetValue("本月.统计富氧量", dr.IsDBNull(5) ? 0 : dr.GetDouble(5));
                        this.SetValue("本月.冷风流量", dr.IsDBNull(6) ? 0 : dr.GetDouble(6));
                        this.SetValue("本月.高压操作时间", dr.IsDBNull(7) ? 0 : dr.GetDouble(7));
                        this.SetValue("本月.平均风温",  dr.IsDBNull(8) ? 0 : dr.GetDouble(8));
                        this.SetValue("本月.炉顶温度",  dr.IsDBNull(9) ? 0 : dr.GetDouble(9));
                        this.SetValue("本月.热风压力",  dr.IsDBNull(10) ? 0 : dr.GetDouble(10));
                        this.SetValue("本月.炉顶压力",  dr.IsDBNull(11) ? 0 : dr.GetDouble(11));
                        this.SetValue("本月.风口损坏数大",  dr.IsDBNull(12) ? 0 : dr.GetDouble(12));
                        this.SetValue("本月.风口损坏数中",  dr.IsDBNull(13) ? 0 : dr.GetDouble(13));
                        this.SetValue("本月.风口损坏数小",  dr.IsDBNull(14) ? 0 : dr.GetDouble(14));
                        this.SetValue("本月.渣口损坏数大",  dr.IsDBNull(15) ? 0 : dr.GetDouble(15));
                        this.SetValue("本月.渣口损坏数中",  dr.IsDBNull(16) ? 0 : dr.GetDouble(16));
                        this.SetValue("本月.渣口损坏数小",  dr.IsDBNull(17) ? 0 : dr.GetDouble(17));
                        this.SetValue("本月.坐料次数",  dr.IsDBNull(18) ? 0 : dr.GetDouble(18));
                        this.SetValue("本月.悬料次数",  dr.IsDBNull(19) ? 0 : dr.GetDouble(19));
                        this.SetValue("本月.崩料次数",  dr.IsDBNull(20) ? 0 : dr.GetDouble(20));
                      //  this.SetValue("本月.干焦粉",  dr.IsDBNull(21) ? 0 : dr.GetDouble(21));
                       this.SetValue(string.Format("本月.废铁", gaolu), dr.IsDBNull(21) ? 0 : dr.GetDouble(21));
                    }

                }
            }
            this.SetValue("本月.CO2除数",CO2chushu);
            this.SetValue("本月.CO除数", COchushu);

            //年消耗
            cmd.CommandText = "select nvl(gaolu,9999),round(avg(CO2),2),round(avg(CO),2),SUM(SHENLIAOPI),SUM(LIAOPI),SUM(FUYANG),SUM(LENGFENGLIULIANG*24),SUM(GAOYASHIJIAN),trunc(avg(fengwen),0),trunc(avg(ludingwendu),0),trunc(avg(refengyali),0),trunc(avg(ludingyali),0),sum(fengkoudatao),sum(fengkouzhongtao),sum(fengkouxiaotao),sum(zhakoudatao),sum(zhakouzhongtao),sum(zhakouxiaotao),sum(zuoliao),sum(xuanliao),sum(bengliao) from XIAOHAO where trunc(RQ,'YEAR')=TRUNC(:rq,'YEAR') and trunc(RQ,'MONTH')<=TRUNC(:rq,'MONTH') group by rollup(gaolu)";
            cmd.Parameters.Clear();
            cmd.Parameters.Add(":rq", OracleType.DateTime).Value = curMonth;
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    if (gaolu != 9999)
                    {
                        double CO2 = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                        double CO = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                        if (CO2 > 0) CO2chushu++;
                        if (CO > 0) COchushu++;

                        this.SetValue(string.Format("累计.高炉{0}.CO2", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                        this.SetValue(string.Format("累计.高炉{0}.CO", gaolu), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                        this.SetValue(string.Format("累计.高炉{0}.深料线", gaolu), dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                        this.SetValue(string.Format("累计.高炉{0}.全部料线", gaolu), dr.IsDBNull(4) ? 0 : dr.GetDouble(4));
                        this.SetValue(string.Format("累计.高炉{0}.富氧流量", gaolu), dr.IsDBNull(5) ? 0 : dr.GetDouble(5));
                        this.SetValue(string.Format("累计.高炉{0}.冷风流量", gaolu), dr.IsDBNull(6) ? 0 : dr.GetDouble(6));
                        this.SetValue(string.Format("累计.高炉{0}.高压操作时间", gaolu), dr.IsDBNull(7) ? 0 : dr.GetDouble(7));
                        this.SetValue(string.Format("累计.高炉{0}.平均风温", gaolu), dr.IsDBNull(8) ? 0 : dr.GetDouble(8));
                        this.SetValue(string.Format("累计.高炉{0}.炉顶温度", gaolu), dr.IsDBNull(9) ? 0 : dr.GetDouble(9));
                        this.SetValue(string.Format("累计.高炉{0}.热风压力", gaolu), dr.IsDBNull(10) ? 0 : dr.GetDouble(10));
                        this.SetValue(string.Format("累计.高炉{0}.炉顶压力", gaolu), dr.IsDBNull(11) ? 0 : dr.GetDouble(11));
                        this.SetValue(string.Format("累计.高炉{0}.风口损坏数大", gaolu), dr.IsDBNull(12) ? 0 : dr.GetDouble(12));
                        this.SetValue(string.Format("累计.高炉{0}.风口损坏数中", gaolu), dr.IsDBNull(13) ? 0 : dr.GetDouble(13));
                        this.SetValue(string.Format("累计.高炉{0}.风口损坏数小", gaolu), dr.IsDBNull(14) ? 0 : dr.GetDouble(14));
                        this.SetValue(string.Format("累计.高炉{0}.渣口损坏数大", gaolu), dr.IsDBNull(15) ? 0 : dr.GetDouble(15));
                        this.SetValue(string.Format("累计.高炉{0}.渣口损坏数中", gaolu), dr.IsDBNull(16) ? 0 : dr.GetDouble(16));
                        this.SetValue(string.Format("累计.高炉{0}.渣口损坏数小", gaolu), dr.IsDBNull(17) ? 0 : dr.GetDouble(17));
                        this.SetValue(string.Format("累计.高炉{0}.坐料次数", gaolu), dr.IsDBNull(18) ? 0 : dr.GetDouble(18));
                        this.SetValue(string.Format("累计.高炉{0}.悬料次数", gaolu), dr.IsDBNull(19) ? 0 : dr.GetDouble(19));
                        this.SetValue(string.Format("累计.高炉{0}.崩料次数", gaolu), dr.IsDBNull(20) ? 0 : dr.GetDouble(20));
                    //    this.SetValue(string.Format("累计.高炉{0}.干焦粉", gaolu), dr.IsDBNull(21) ? 0 : dr.GetDouble(21));
                    }
                    else
                    {
                        this.SetValue("累计.CO2", dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                        this.SetValue("累计.CO", dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                        this.SetValue("累计.深料线", dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                        this.SetValue("累计.全部料线", dr.IsDBNull(4) ? 0 : dr.GetDouble(4));
                        this.SetValue("累计.富氧流量", dr.IsDBNull(5) ? 0 : dr.GetDouble(5));
                        this.SetValue("累计.冷风流量", dr.IsDBNull(6) ? 0 : dr.GetDouble(6));
                        this.SetValue("累计.高压操作时间", dr.IsDBNull(7) ? 0 : dr.GetDouble(7));
                        this.SetValue("累计.平均风温", dr.IsDBNull(8) ? 0 : dr.GetDouble(8));
                        this.SetValue("累计.炉顶温度", dr.IsDBNull(9) ? 0 : dr.GetDouble(9));
                        this.SetValue("累计.热风压力", dr.IsDBNull(10) ? 0 : dr.GetDouble(10));
                        this.SetValue("累计.炉顶压力", dr.IsDBNull(11) ? 0 : dr.GetDouble(11));
                        this.SetValue("累计.风口损坏数大", dr.IsDBNull(12) ? 0 : dr.GetDouble(12));
                        this.SetValue("累计.风口损坏数中", dr.IsDBNull(13) ? 0 : dr.GetDouble(13));
                        this.SetValue("累计.风口损坏数小", dr.IsDBNull(14) ? 0 : dr.GetDouble(14));
                        this.SetValue("累计.渣口损坏数大", dr.IsDBNull(15) ? 0 : dr.GetDouble(15));
                        this.SetValue("累计.渣口损坏数中", dr.IsDBNull(16) ? 0 : dr.GetDouble(16));
                        this.SetValue("累计.渣口损坏数小", dr.IsDBNull(17) ? 0 : dr.GetDouble(17));
                        this.SetValue("累计.坐料次数", dr.IsDBNull(18) ? 0 : dr.GetDouble(18));
                        this.SetValue("累计.悬料次数", dr.IsDBNull(19) ? 0 : dr.GetDouble(19));
                        this.SetValue("累计.崩料次数", dr.IsDBNull(20) ? 0 : dr.GetDouble(20));
                     //   this.SetValue("累计.干焦粉", dr.IsDBNull(21) ? 0 : dr.GetDouble(21));
                    }

                }
            }
            //年工艺称
            cmd.CommandText = "select 高炉,sum(工艺称) from 原料消耗 where trunc(日期,'YEAR')=TRUNC(:rq,'YEAR') and trunc(日期,'MONTH')<=TRUNC(:rq,'MONTH')  group by 高炉";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("累计.高炉{0}.工艺称", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();

            //制钢铁、铸造铁SI标准偏差
            cmd.CommandText = "select nvl(gaolu,9999),case when fesi<=1.25 then '制钢铁' else '铸造铁' end,round(stddev(FESI),2) from ddluci where zdsj>=:rq1 and zdsj<:rq2  and dksj is not null group by rollup(gaolu),case when fesi<=1.25 then '制钢铁' else '铸造铁' end";
            cmd.Parameters.Clear();
            cmd.Parameters.Add(":rq1", OracleType.DateTime).Value = curYear;
            cmd.Parameters.Add(":rq2", OracleType.DateTime).Value = lastDay.AddDays(1);
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    if (gaolu != 9999)
                    {
                        this.SetValue(string.Format("累计.高炉{0}.{1}{2}", gaolu, dr.GetString(1), "Si偏差"), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                    }
                    else
                    {
                        this.SetValue(string.Format("累计.{0}{1}", dr.GetString(1), "Si偏差"), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                    }
                }
            }
            dr.Close();

            //调整参数
            cmd.Parameters.Clear();
            cmd.Parameters.Add(":nian", OracleType.Int32).Value = nian;
            cmd.Parameters.Add(":yue", OracleType.Int32).Value = yue;
            cmd.CommandText = "select 名称,重量,TFE,CAO from 每月矿 where 年=:nian and 月=:yue";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                string mc = dr.IsDBNull(0) ? "" : dr.GetString(0).Replace("竖炉球", "竖球");
                if (mc != "")
                {
                    this.SetValue(string.Format("本月.调整{0}", mc), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                    this.SetValue(string.Format("本月.{0}品位", mc), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                    this.SetValue(string.Format("本月.{0}CaO", mc), dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                }
            }
            dr.Close();

            cmd.CommandText = "select 干毛焦,湿焦粉,焦丁,焦炭灰份 from 每月焦 where 年=:nian and 月=:yue";
            dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                this.SetValue("本月.调整干毛焦", dr.IsDBNull(0) ? 0 : dr.GetDouble(0));
                this.SetValue("本月.调整干焦粉", dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                this.SetValue("本月.调整焦丁", dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                this.SetValue("本月.焦炭灰份", dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
            }
            dr.Close();

            cmd.CommandText = "select 重量*(1-水份/100),灰份 from 每月煤 where 年=:nian and 月=:yue";
            dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                this.SetValue("本月.调整煤粉", dr.IsDBNull(0) ? 0 : dr.GetDouble(0));
                this.SetValue("本月.煤粉灰份", dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
            }
            dr.Close();

            cmd.CommandText = "select 富氧量 from 每月能耗 where 年=:nian and 月=:yue";
            dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                this.SetValue("本月.调整富氧量", dr.IsDBNull(0) ? 0 : dr.GetDouble(0));
            }
            dr.Close();



            if (yue > 1)
            {
                StringBuilder sqlSb = new StringBuilder();
                sqlSb.Append("select P01单位,P03全铁产量,P04合格铁,P05制钢铁,P07折合产量,一级铁,优质铁,P61干毛焦,P62干焦粉,P63入炉焦炭总耗,P65煤粉总耗,P66焦丁总耗");
                sqlSb.Append(",P50机烧消耗,P51竖炉球消耗,P53其它熟料消耗,P54本溪矿消耗,P55其它生料消耗,P56废铁总耗,瓦斯灰,P60七号称");
                sqlSb.Append(",P40风口大套损坏数,P41风口中套损坏数,P42风口小套损坏数,P43渣口大套损坏数,P44渣口中套损坏数,P45渣口小套损坏数");
                sqlSb.Append(",P37悬料次数,P38坐料次数,P39崩料次数,P20休风时间,P22慢风时间,P29冷风流量,富氧量");
                sqlSb.Append(",收入含铁,支出含铁,高压操作时间,有效工作时间,理论渣量,深料线,全部料线,高炉有效容积,高炉实际容积 ");
                sqlSb.Append(",P30平均风温,P32热风压力,P34炉顶温度,P33炉顶压力");
                sqlSb.Append(",P24煤气成分CO2,P25煤气成分计算1,入炉矿含铁 ");
                sqlSb.Append(",PB块,纽曼块,钛球,锰矿,硅石,白云石,蛇纹石,萤石,球团矿,国内球团矿,进口球团矿,其它块矿,高钛球团矿,高品位钛球,其它熔剂 ");
                sqlSb.Append(" from 技术月报 where P年=:nian and P月=:yue and P02项目='累计'");
                cmd.CommandText = sqlSb.ToString();
                cmd.Parameters.Clear();


                //this.SetValue(string.Format("上月.高炉{0}.萤石", gaolu), shuliangDr.IsDBNull(56) ? 0 : shuliangDr.GetDouble(56));
                //this.SetValue(string.Format("上月.高炉{0}.球团矿", gaolu), shuliangDr.IsDBNull(57) ? 0 : shuliangDr.GetDouble(57));
                //this.SetValue(string.Format("上月.高炉{0}.国内球团矿", gaolu), shuliangDr.IsDBNull(58) ? 0 : shuliangDr.GetDouble(58));
                //this.SetValue(string.Format("上月.高炉{0}.进口球团矿", gaolu), shuliangDr.IsDBNull(59) ? 0 : shuliangDr.GetDouble(59));
                //this.SetValue(string.Format("上月.高炉{0}.其它块矿", gaolu), shuliangDr.IsDBNull(60) ? 0 : shuliangDr.GetDouble(60));
                //this.SetValue(string.Format("上月.高炉{0}.高钛球团矿", gaolu), shuliangDr.IsDBNull(61) ? 0 : shuliangDr.GetDouble(61));
                //this.SetValue(string.Format("上月.高炉{0}.高品位钛球", gaolu), shuliangDr.IsDBNull(62) ? 0 : shuliangDr.GetDouble(62));
                //this.SetValue(string.Format("上月.高炉{0}.其它熔剂", gaolu), shuliangDr.IsDBNull(63) ? 0 : shuliangDr.GetDouble(63));
                cmd.Parameters.Add(":nian", OracleType.Int32).Value = nian;
                cmd.Parameters.Add(":yue", OracleType.Int32).Value = yue - 1;
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string gaolustr = dr.GetString(0);
                    if (gaolustr.Contains("#") && gaolustr.Length > 1)
                    {
                        int gaolu;
                        if (int.TryParse(gaolustr.Substring(0, gaolustr.Length - 1), out gaolu))
                        {
                            this.SetValue(string.Format("上月累计.高炉{0}.全铁产量", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                            this.SetValue(string.Format("上月累计.高炉{0}.合格铁", gaolu), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                            this.SetValue(string.Format("上月累计.高炉{0}.制钢铁", gaolu), dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                            this.SetValue(string.Format("上月累计.高炉{0}.折算产量", gaolu), dr.IsDBNull(4) ? 0 : dr.GetDouble(4));
                            this.SetValue(string.Format("上月累计.高炉{0}.一级铁", gaolu), dr.IsDBNull(5) ? 0 : dr.GetDouble(5));
                            this.SetValue(string.Format("上月累计.高炉{0}.优质铁", gaolu), dr.IsDBNull(6) ? 0 : dr.GetDouble(6));
                            this.SetValue(string.Format("上月累计.高炉{0}.干毛焦", gaolu), dr.IsDBNull(7) ? 0 : dr.GetDouble(7));
                            this.SetValue(string.Format("上月累计.高炉{0}.干焦粉", gaolu), dr.IsDBNull(8) ? 0 : dr.GetDouble(8));
                            this.SetValue(string.Format("上月累计.高炉{0}.入炉焦炭", gaolu), dr.IsDBNull(9) ? 0 : dr.GetDouble(9));
                            this.SetValue(string.Format("上月累计.高炉{0}.煤粉", gaolu), dr.IsDBNull(10) ? 0 : dr.GetDouble(10));
                            this.SetValue(string.Format("上月累计.高炉{0}.焦丁", gaolu), dr.IsDBNull(11) ? 0 : dr.GetDouble(11));
                            this.SetValue(string.Format("上月累计.高炉{0}.机烧", gaolu), dr.IsDBNull(12) ? 0 : dr.GetDouble(12));
                            this.SetValue(string.Format("上月累计.高炉{0}.竖球", gaolu), dr.IsDBNull(13) ? 0 : dr.GetDouble(13));
                            this.SetValue(string.Format("上月累计.高炉{0}.其它熟料", gaolu), dr.IsDBNull(14) ? 0 : dr.GetDouble(14));
                            this.SetValue(string.Format("上月累计.高炉{0}.本溪矿", gaolu), dr.IsDBNull(15) ? 0 : dr.GetDouble(15));
                            this.SetValue(string.Format("上月累计.高炉{0}.其它生料", gaolu), dr.IsDBNull(16) ? 0 : dr.GetDouble(16));
                            this.SetValue(string.Format("上月累计.高炉{0}.废铁", gaolu), dr.IsDBNull(17) ? 0 : dr.GetDouble(17));
                            this.SetValue(string.Format("上月累计.高炉{0}.瓦斯灰", gaolu), dr.IsDBNull(18) ? 0 : dr.GetDouble(18));
                            this.SetValue(string.Format("上月累计.高炉{0}.七号称", gaolu), dr.IsDBNull(19) ? 0 : dr.GetDouble(19));
                            this.SetValue(string.Format("上月累计.高炉{0}.风口损坏数大", gaolu), dr.IsDBNull(20) ? 0 : dr.GetDouble(20));
                            this.SetValue(string.Format("上月累计.高炉{0}.风口损坏数中", gaolu), dr.IsDBNull(21) ? 0 : dr.GetDouble(21));
                            this.SetValue(string.Format("上月累计.高炉{0}.风口损坏数小", gaolu), dr.IsDBNull(22) ? 0 : dr.GetDouble(22));
                            this.SetValue(string.Format("上月累计.高炉{0}.渣口损坏数大", gaolu), dr.IsDBNull(23) ? 0 : dr.GetDouble(23));
                            this.SetValue(string.Format("上月累计.高炉{0}.渣口损坏数中", gaolu), dr.IsDBNull(24) ? 0 : dr.GetDouble(24));
                            this.SetValue(string.Format("上月累计.高炉{0}.渣口损坏数小", gaolu), dr.IsDBNull(25) ? 0 : dr.GetDouble(25));
                            this.SetValue(string.Format("上月累计.高炉{0}.悬料次数", gaolu), dr.IsDBNull(26) ? 0 : dr.GetDouble(26));
                            this.SetValue(string.Format("上月累计.高炉{0}.坐料次数", gaolu), dr.IsDBNull(27) ? 0 : dr.GetDouble(27));
                            this.SetValue(string.Format("上月累计.高炉{0}.崩料次数", gaolu), dr.IsDBNull(28) ? 0 : dr.GetDouble(28));
                            this.SetValue(string.Format("上月累计.高炉{0}.休风时间", gaolu), dr.IsDBNull(29) ? 0 : dr.GetDouble(29));
                            this.SetValue(string.Format("上月累计.高炉{0}.慢风时间", gaolu), dr.IsDBNull(30) ? 0 : dr.GetDouble(30));
                            this.SetValue(string.Format("上月累计.高炉{0}.冷风流量", gaolu), dr.IsDBNull(31) ? 0 : dr.GetDouble(31));
                            this.SetValue(string.Format("上月累计.高炉{0}.富氧量", gaolu), dr.IsDBNull(32) ? 0 : dr.GetDouble(32));
                            this.SetValue(string.Format("上月累计.高炉{0}.收入含铁", gaolu), dr.IsDBNull(33) ? 0 : dr.GetDouble(33));
                            this.SetValue(string.Format("上月累计.高炉{0}.支出含铁", gaolu), dr.IsDBNull(34) ? 0 : dr.GetDouble(34));
                            this.SetValue(string.Format("上月累计.高炉{0}.高压操作时间", gaolu), dr.IsDBNull(35) ? 0 : dr.GetDouble(35));
                            this.SetValue(string.Format("上月累计.高炉{0}.有效工作时间", gaolu), dr.IsDBNull(36) ? 0 : dr.GetDouble(36));
                            this.SetValue(string.Format("上月累计.高炉{0}.理论渣量", gaolu), dr.IsDBNull(37) ? 0 : dr.GetDouble(37));
                            this.SetValue(string.Format("上月累计.高炉{0}.深料线", gaolu), dr.IsDBNull(38) ? 0 : dr.GetDouble(38));
                            this.SetValue(string.Format("上月累计.高炉{0}.全部料线", gaolu), dr.IsDBNull(39) ? 0 : dr.GetDouble(39));
                            this.SetValue(string.Format("上月累计.高炉{0}.高炉有效容积", gaolu), dr.IsDBNull(40) ? 0 : dr.GetDouble(40));
                            this.SetValue(string.Format("上月累计.高炉{0}.高炉实际容积", gaolu), dr.IsDBNull(41) ? 0 : dr.GetDouble(41));
                            this.SetValue(string.Format("上月累计.高炉{0}.平均风温", gaolu), dr.IsDBNull(42) ? 0 : dr.GetDouble(42));
                            this.SetValue(string.Format("上月累计.高炉{0}.热风压力", gaolu), dr.IsDBNull(43) ? 0 : dr.GetDouble(43));
                            this.SetValue(string.Format("上月累计.高炉{0}.炉顶温度", gaolu), dr.IsDBNull(44) ? 0 : dr.GetDouble(44));
                            this.SetValue(string.Format("上月累计.高炉{0}.炉顶压力", gaolu), dr.IsDBNull(45) ? 0 : dr.GetDouble(45));
                            this.SetValue(string.Format("上月累计.高炉{0}.CO2", gaolu), dr.IsDBNull(46) ? 0 : dr.GetDouble(46));
                            this.SetValue(string.Format("上月累计.高炉{0}.CO2率", gaolu), dr.IsDBNull(47) ? 0 : dr.GetDouble(47));
                            this.SetValue(string.Format("上月累计.高炉{0}.入炉矿含铁", gaolu), dr.IsDBNull(48) ? 0 : dr.GetDouble(48));
                            this.SetValue(string.Format("上月累计.高炉{0}.PB块", gaolu), dr.IsDBNull(49) ? 0 : dr.GetDouble(49));
                            this.SetValue(string.Format("上月累计.高炉{0}.纽曼块", gaolu), dr.IsDBNull(50) ? 0 : dr.GetDouble(50));
                            this.SetValue(string.Format("上月累计.高炉{0}.钛球", gaolu), dr.IsDBNull(51) ? 0 : dr.GetDouble(51));
                            this.SetValue(string.Format("上月累计.高炉{0}.锰矿", gaolu), dr.IsDBNull(52) ? 0 : dr.GetDouble(52));
                            this.SetValue(string.Format("上月累计.高炉{0}.硅石", gaolu), dr.IsDBNull(53) ? 0 : dr.GetDouble(53));
                            this.SetValue(string.Format("上月累计.高炉{0}.白云石", gaolu), dr.IsDBNull(54) ? 0 : dr.GetDouble(54));
                            this.SetValue(string.Format("上月累计.高炉{0}.蛇纹石", gaolu), dr.IsDBNull(55) ? 0 : dr.GetDouble(55));

                            this.SetValue(string.Format("上月累计.高炉{0}.萤石", gaolu), dr.IsDBNull(56) ? 0 : dr.GetDouble(56));
                            this.SetValue(string.Format("上月累计.高炉{0}.球团矿", gaolu), dr.IsDBNull(57) ? 0 : dr.GetDouble(57));
                            this.SetValue(string.Format("上月累计.高炉{0}.国内球团矿", gaolu), dr.IsDBNull(58) ? 0 : dr.GetDouble(58));
                            this.SetValue(string.Format("上月累计.高炉{0}.进口球团矿", gaolu), dr.IsDBNull(59) ? 0 : dr.GetDouble(59));
                            this.SetValue(string.Format("上月累计.高炉{0}.其它块矿", gaolu), dr.IsDBNull(60) ? 0 : dr.GetDouble(60));
                            this.SetValue(string.Format("上月累计.高炉{0}.高钛球团矿", gaolu), dr.IsDBNull(61) ? 0 : dr.GetDouble(61));
                            this.SetValue(string.Format("上月累计.高炉{0}.高品位钛球", gaolu), dr.IsDBNull(62) ? 0 : dr.GetDouble(62));
                            this.SetValue(string.Format("上月累计.高炉{0}.其它熔剂", gaolu), dr.IsDBNull(63) ? 0 : dr.GetDouble(63));
                 
                        }
                    }
                    else if (gaolustr == "Q")
                    {
                        this.SetValue("上月累计.全铁产量", dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                        this.SetValue("上月累计.合格铁", dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                        this.SetValue("上月累计.制钢铁", dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                        this.SetValue("上月累计.折算产量", dr.IsDBNull(4) ? 0 : dr.GetDouble(4));
                        this.SetValue("上月累计.一级铁", dr.IsDBNull(5) ? 0 : dr.GetDouble(5));
                        this.SetValue("上月累计.优质铁", dr.IsDBNull(6) ? 0 : dr.GetDouble(6));
                        this.SetValue("上月累计.干毛焦", dr.IsDBNull(7) ? 0 : dr.GetDouble(7));
                        this.SetValue("上月累计.干焦粉", dr.IsDBNull(8) ? 0 : dr.GetDouble(8));
                        this.SetValue("上月累计.入炉焦炭", dr.IsDBNull(9) ? 0 : dr.GetDouble(9));
                        this.SetValue("上月累计.煤粉", dr.IsDBNull(10) ? 0 : dr.GetDouble(10));
                        this.SetValue("上月累计.焦丁", dr.IsDBNull(11) ? 0 : dr.GetDouble(11));
                        this.SetValue("上月累计.机烧", dr.IsDBNull(12) ? 0 : dr.GetDouble(12));
                        this.SetValue("上月累计.竖球", dr.IsDBNull(13) ? 0 : dr.GetDouble(13));
                        this.SetValue("上月累计.其它熟料", dr.IsDBNull(14) ? 0 : dr.GetDouble(14));
                        this.SetValue("上月累计.本溪矿", dr.IsDBNull(15) ? 0 : dr.GetDouble(15));
                        this.SetValue("上月累计.其它生料", dr.IsDBNull(16) ? 0 : dr.GetDouble(16));
                        this.SetValue("上月累计.废铁", dr.IsDBNull(17) ? 0 : dr.GetDouble(17));
                        this.SetValue("上月累计.瓦斯灰", dr.IsDBNull(18) ? 0 : dr.GetDouble(18));
                        this.SetValue("上月累计.七号称", dr.IsDBNull(19) ? 0 : dr.GetDouble(19));
                        this.SetValue("上月累计.风口损坏数大", dr.IsDBNull(20) ? 0 : dr.GetDouble(20));
                        this.SetValue("上月累计.风口损坏数中", dr.IsDBNull(21) ? 0 : dr.GetDouble(21));
                        this.SetValue("上月累计.风口损坏数小", dr.IsDBNull(22) ? 0 : dr.GetDouble(22));
                        this.SetValue("上月累计.渣口损坏数大", dr.IsDBNull(23) ? 0 : dr.GetDouble(23));
                        this.SetValue("上月累计.渣口损坏数中", dr.IsDBNull(24) ? 0 : dr.GetDouble(24));
                        this.SetValue("上月累计.渣口损坏数小", dr.IsDBNull(25) ? 0 : dr.GetDouble(25));
                        this.SetValue("上月累计.悬料次数", dr.IsDBNull(26) ? 0 : dr.GetDouble(26));
                        this.SetValue("上月累计.坐料次数", dr.IsDBNull(27) ? 0 : dr.GetDouble(27));
                        this.SetValue("上月累计.崩料次数", dr.IsDBNull(28) ? 0 : dr.GetDouble(28));
                        this.SetValue("上月累计.休风时间", dr.IsDBNull(29) ? 0 : dr.GetDouble(29));
                        this.SetValue("上月累计.慢风时间", dr.IsDBNull(30) ? 0 : dr.GetDouble(30));
                        this.SetValue("上月累计.冷风流量", dr.IsDBNull(31) ? 0 : dr.GetDouble(31));
                        this.SetValue("上月累计.富氧量", dr.IsDBNull(32) ? 0 : dr.GetDouble(32));
                        this.SetValue("上月累计.收入含铁", dr.IsDBNull(33) ? 0 : dr.GetDouble(33));
                        this.SetValue("上月累计.支出含铁", dr.IsDBNull(34) ? 0 : dr.GetDouble(34));
                        this.SetValue("上月累计.高压操作时间", dr.IsDBNull(35) ? 0 : dr.GetDouble(35));
                        this.SetValue("上月累计.有效工作时间", dr.IsDBNull(36) ? 0 : dr.GetDouble(36));
                        this.SetValue("上月累计.理论渣量", dr.IsDBNull(37) ? 0 : dr.GetDouble(37));
                        this.SetValue("上月累计.深料线", dr.IsDBNull(38) ? 0 : dr.GetDouble(38));
                        this.SetValue("上月累计.全部料线", dr.IsDBNull(39) ? 0 : dr.GetDouble(39));
                        this.SetValue("上月累计.高炉有效容积", dr.IsDBNull(40) ? 0 : dr.GetDouble(40));
                        this.SetValue("上月累计.高炉实际容积", dr.IsDBNull(41) ? 0 : dr.GetDouble(41));
                        this.SetValue("上月累计.平均风温", dr.IsDBNull(42) ? 0 : dr.GetDouble(42));
                        this.SetValue("上月累计.热风压力", dr.IsDBNull(43) ? 0 : dr.GetDouble(43));
                        this.SetValue("上月累计.炉顶温度", dr.IsDBNull(44) ? 0 : dr.GetDouble(44));
                        this.SetValue("上月累计.炉顶压力", dr.IsDBNull(45) ? 0 : dr.GetDouble(45));
                        this.SetValue("上月累计.CO2", dr.IsDBNull(46) ? 0 : dr.GetDouble(46));
                        this.SetValue("上月累计.CO2率", dr.IsDBNull(47) ? 0 : dr.GetDouble(47));
                        this.SetValue("上月累计.入炉矿含铁", dr.IsDBNull(48) ? 0 : dr.GetDouble(48));
                        this.SetValue("上月累计.PB块", dr.IsDBNull(49) ? 0 : dr.GetDouble(49));
                        this.SetValue("上月累计.纽曼块", dr.IsDBNull(50) ? 0 : dr.GetDouble(50));
                        this.SetValue("上月累计.钛球", dr.IsDBNull(51) ? 0 : dr.GetDouble(51));
                        this.SetValue("上月累计.锰矿", dr.IsDBNull(52) ? 0 : dr.GetDouble(52));
                        this.SetValue("上月累计.硅石", dr.IsDBNull(53) ? 0 : dr.GetDouble(53));
                        this.SetValue("上月累计.白云石", dr.IsDBNull(54) ? 0 : dr.GetDouble(54));
                        this.SetValue("上月累计.蛇纹石", dr.IsDBNull(55) ? 0 : dr.GetDouble(55));
                        this.SetValue("上月累计.萤石", dr.IsDBNull(56) ? 0 : dr.GetDouble(56));
                        this.SetValue("上月累计.球团矿", dr.IsDBNull(57) ? 0 : dr.GetDouble(57));
                        this.SetValue("上月累计.国内球团矿", dr.IsDBNull(58) ? 0 : dr.GetDouble(58));
                        this.SetValue("上月累计.进口球团矿", dr.IsDBNull(59) ? 0 : dr.GetDouble(59));
                        this.SetValue("上月累计.其它块矿", dr.IsDBNull(60) ? 0 : dr.GetDouble(60));
                        this.SetValue("上月累计.高钛球团矿", dr.IsDBNull(61) ? 0 : dr.GetDouble(61));
                        this.SetValue("上月累计.高品位钛球", dr.IsDBNull(62) ? 0 : dr.GetDouble(62));
                        this.SetValue("上月累计.其它熔剂", dr.IsDBNull(63) ? 0 : dr.GetDouble(63));
                    }
                }
                dr.Close();
            }

            #region  成分
            //铁、渣成分
            StringBuilder chengFenSql = new StringBuilder();
            chengFenSql.Append("select nvl(gaolu,9999)");
            chengFenSql.Append(",round(sum(feliang*FESI)/sum(feliang),2),round(sum(feliang*FEMn)/sum(feliang),2),round(sum(feliang*FEP)/sum(feliang),3),round(sum(feliang*FES)/sum(feliang),3),round(sum(feliang*(100-NVL(FEC,0)-NVL(FESI,0)-NVL(FEMN,0)-NVL(FEP,0)-NVL(FES,0)-NVL(FETI,0)))/sum(feliang),2)");
            chengFenSql.Append(",round(avg(ZHACAO),2),round(avg(ZHASIO2),2),round(avg(ZHAAL2O3),2),round(avg(ZHAMGO),2),round(avg(ZHAS),2),round(avg(ZHAR2),2)");
            chengFenSql.Append(" from ddluci");
            chengFenSql.Append(" where trunc(zdsj,'MONTH')=:RQ  and dksj is not null group by rollup(gaolu)");

            OracleCommand chengFenCmd = new OracleCommand(chengFenSql.ToString(),conn);
            chengFenCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = curMonth;

            OracleDataReader chengFenDr = chengFenCmd.ExecuteReader();
            while (chengFenDr.Read())
            {
                if (!chengFenDr.IsDBNull(0))
                {
                    int gaolu = chengFenDr.GetInt32(0);
                    if (gaolu != 9999)
                    {
                        this.SetValue(string.Format("本月.高炉{0}.{1}", gaolu, "Fe_Si"), chengFenDr.IsDBNull(1) ? 0 : chengFenDr.GetDouble(1));
                        this.SetValue(string.Format("本月.高炉{0}.{1}", gaolu, "Fe_Mn"), chengFenDr.IsDBNull(2) ? 0 : chengFenDr.GetDouble(2));
                        this.SetValue(string.Format("本月.高炉{0}.{1}", gaolu, "Fe_P"), chengFenDr.IsDBNull(3) ? 0 : chengFenDr.GetDouble(3));
                        this.SetValue(string.Format("本月.高炉{0}.{1}", gaolu, "Fe_S"), chengFenDr.IsDBNull(4) ? 0 : chengFenDr.GetDouble(4));
                        this.SetValue(string.Format("本月.高炉{0}.{1}", gaolu, "Fe_Fe"), chengFenDr.IsDBNull(5) ? 0 : chengFenDr.GetDouble(5));
                        this.SetValue(string.Format("本月.高炉{0}.{1}", gaolu, "Zha_CaO"), chengFenDr.IsDBNull(6) ? 0 : chengFenDr.GetDouble(6));
                        this.SetValue(string.Format("本月.高炉{0}.{1}", gaolu, "Zha_SiO2"), chengFenDr.IsDBNull(7) ? 0 : chengFenDr.GetDouble(7));
                        this.SetValue(string.Format("本月.高炉{0}.{1}", gaolu, "Zha_Al2O3"), chengFenDr.IsDBNull(8) ? 0 : chengFenDr.GetDouble(8));
                        this.SetValue(string.Format("本月.高炉{0}.{1}", gaolu, "Zha_MgO"), chengFenDr.IsDBNull(9) ? 0 : chengFenDr.GetDouble(9));
                        this.SetValue(string.Format("本月.高炉{0}.{1}", gaolu, "Zha_S"), chengFenDr.IsDBNull(10) ? 0 : chengFenDr.GetDouble(10));
                        this.SetValue(string.Format("本月.高炉{0}.{1}", gaolu, "碱度"), chengFenDr.IsDBNull(11) ? 0 : chengFenDr.GetDouble(11));
                    }
                    else
                    {
                        this.SetValue("本月.Fe_Si", chengFenDr.IsDBNull(1) ? 0 : chengFenDr.GetDouble(1));
                        this.SetValue("本月.Fe_Mn", chengFenDr.IsDBNull(2) ? 0 : chengFenDr.GetDouble(2));
                        this.SetValue("本月.Fe_P", chengFenDr.IsDBNull(3) ? 0 : chengFenDr.GetDouble(3));
                        this.SetValue("本月.Fe_S", chengFenDr.IsDBNull(4) ? 0 : chengFenDr.GetDouble(4));
                        this.SetValue("本月.Fe_Fe", chengFenDr.IsDBNull(5) ? 0 : chengFenDr.GetDouble(5));
                        this.SetValue("本月.Zha_CaO", chengFenDr.IsDBNull(6) ? 0 : chengFenDr.GetDouble(6));
                        this.SetValue("本月.Zha_SiO2", chengFenDr.IsDBNull(7) ? 0 : chengFenDr.GetDouble(7));
                        this.SetValue("本月.Zha_Al2O3", chengFenDr.IsDBNull(8) ? 0 : chengFenDr.GetDouble(8));
                        this.SetValue("本月.Zha_MgO", chengFenDr.IsDBNull(9) ? 0 : chengFenDr.GetDouble(9));
                        this.SetValue("本月.Zha_S", chengFenDr.IsDBNull(10) ? 0 : chengFenDr.GetDouble(10));
                        this.SetValue("本月.碱度", chengFenDr.IsDBNull(11) ? 0 : chengFenDr.GetDouble(11));
                    }
                }
            }
            chengFenDr.Close();

            //累计铁、渣成分
            StringBuilder chengFenLeiJiSql = new StringBuilder();
            chengFenLeiJiSql.Append("select nvl(gaolu,9999)");
            chengFenLeiJiSql.Append(",round(sum(feliang*FESI)/sum(feliang),2),round(sum(feliang*FEMn)/sum(feliang),2),round(sum(feliang*FEP)/sum(feliang),3),round(sum(feliang*FES)/sum(feliang),3),round(sum(feliang*(100-NVL(FEC,0)-NVL(FESI,0)-NVL(FEMN,0)-NVL(FEP,0)-NVL(FES,0)-NVL(FETI,0)))/sum(feliang),2)");
            chengFenLeiJiSql.Append(",round(avg(ZHACAO),2),round(avg(ZHASIO2),2),round(avg(ZHAAL2O3),2),round(avg(ZHAMGO),2),round(avg(ZHAS),2),round(avg(ZHAR2),2)");
            chengFenLeiJiSql.Append(" from ddluci");
            chengFenLeiJiSql.Append(" where zdsj>=:RQ1 and zdsj<:RQ2  and dksj is not null group by rollup(gaolu)");

            OracleCommand chengFenLeiJiCmd = new OracleCommand(chengFenLeiJiSql.ToString(),conn);
            chengFenLeiJiCmd.Parameters.Add(":RQ1", OracleType.DateTime).Value = curYear;
            chengFenLeiJiCmd.Parameters.Add(":RQ2", OracleType.DateTime).Value = lastDay.AddDays(1);

            OracleDataReader chengFenLeiJiDr = chengFenLeiJiCmd.ExecuteReader();
            while (chengFenLeiJiDr.Read())
            {
                if (!chengFenLeiJiDr.IsDBNull(0))
                {
                    int gaolu = chengFenLeiJiDr.GetInt32(0);
                    if (gaolu != 9999)
                    {
                        this.SetValue(string.Format("累计.高炉{0}.{1}", gaolu, "Fe_Si"), chengFenLeiJiDr.IsDBNull(1) ? 0 : chengFenLeiJiDr.GetDouble(1));
                        this.SetValue(string.Format("累计.高炉{0}.{1}", gaolu, "Fe_Mn"), chengFenLeiJiDr.IsDBNull(2) ? 0 : chengFenLeiJiDr.GetDouble(2));
                        this.SetValue(string.Format("累计.高炉{0}.{1}", gaolu, "Fe_P"), chengFenLeiJiDr.IsDBNull(3) ? 0 : chengFenLeiJiDr.GetDouble(3));
                        this.SetValue(string.Format("累计.高炉{0}.{1}", gaolu, "Fe_S"), chengFenLeiJiDr.IsDBNull(4) ? 0 : chengFenLeiJiDr.GetDouble(4));
                        this.SetValue(string.Format("累计.高炉{0}.{1}", gaolu, "Fe_Fe"), chengFenLeiJiDr.IsDBNull(5) ? 0 : chengFenLeiJiDr.GetDouble(5));
                        this.SetValue(string.Format("累计.高炉{0}.{1}", gaolu, "Zha_CaO"), chengFenLeiJiDr.IsDBNull(6) ? 0 : chengFenLeiJiDr.GetDouble(6));
                        this.SetValue(string.Format("累计.高炉{0}.{1}", gaolu, "Zha_SiO2"), chengFenLeiJiDr.IsDBNull(7) ? 0 : chengFenLeiJiDr.GetDouble(7));
                        this.SetValue(string.Format("累计.高炉{0}.{1}", gaolu, "Zha_Al2O3"), chengFenLeiJiDr.IsDBNull(8) ? 0 : chengFenLeiJiDr.GetDouble(8));
                        this.SetValue(string.Format("累计.高炉{0}.{1}", gaolu, "Zha_MgO"), chengFenLeiJiDr.IsDBNull(9) ? 0 : chengFenLeiJiDr.GetDouble(9));
                        this.SetValue(string.Format("累计.高炉{0}.{1}", gaolu, "Zha_S"), chengFenLeiJiDr.IsDBNull(10) ? 0 : chengFenLeiJiDr.GetDouble(10));
                        this.SetValue(string.Format("累计.高炉{0}.{1}", gaolu, "碱度"), chengFenLeiJiDr.IsDBNull(11) ? 0 : chengFenLeiJiDr.GetDouble(11));
                    }
                    else
                    {
                        this.SetValue("累计.Fe_Si", chengFenLeiJiDr.IsDBNull(1) ? 0 : chengFenLeiJiDr.GetDouble(1));
                        this.SetValue("累计.Fe_Mn", chengFenLeiJiDr.IsDBNull(2) ? 0 : chengFenLeiJiDr.GetDouble(2));
                        this.SetValue("累计.Fe_P", chengFenLeiJiDr.IsDBNull(3) ? 0 : chengFenLeiJiDr.GetDouble(3));
                        this.SetValue("累计.Fe_S", chengFenLeiJiDr.IsDBNull(4) ? 0 : chengFenLeiJiDr.GetDouble(4));
                        this.SetValue("累计.Fe_Fe", chengFenLeiJiDr.IsDBNull(5) ? 0 : chengFenLeiJiDr.GetDouble(5));
                        this.SetValue("累计.Zha_CaO", chengFenLeiJiDr.IsDBNull(6) ? 0 : chengFenLeiJiDr.GetDouble(6));
                        this.SetValue("累计.Zha_SiO2", chengFenLeiJiDr.IsDBNull(7) ? 0 : chengFenLeiJiDr.GetDouble(7));
                        this.SetValue("累计.Zha_Al2O3", chengFenLeiJiDr.IsDBNull(8) ? 0 : chengFenLeiJiDr.GetDouble(8));
                        this.SetValue("累计.Zha_MgO", chengFenLeiJiDr.IsDBNull(9) ? 0 : chengFenLeiJiDr.GetDouble(9));
                        this.SetValue("累计.Zha_S", chengFenLeiJiDr.IsDBNull(10) ? 0 : chengFenLeiJiDr.GetDouble(10));
                        this.SetValue("累计.碱度", chengFenLeiJiDr.IsDBNull(11) ? 0 : chengFenLeiJiDr.GetDouble(11));
                    }
                }
            }
            chengFenLeiJiDr.Close();


            #endregion 

            conn.Close();

        }
    }
}
