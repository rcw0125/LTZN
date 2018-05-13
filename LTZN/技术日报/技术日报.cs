using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;
using LTZN.CalFramework;
using LTZN.Repository;

namespace LTZN.技术日报
{
    public class 技术日报:QueryModel
    {
        public 技术日报()
            : base("技术日报")
        {

        }

        public void getDataBy(DateTime rq)
        {
            conn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn; 
            cmd.Parameters.Add(":rq", OracleType.DateTime).Value = rq;

            #region 本日部分
            //炼钢铁
            cmd.CommandText = "select gaolu,case when fesi<=1.25 then '炼钢铁' else '铸造铁' end,sum(feliang) from ddluci where trunc(zdsj)=:rq and dksj is not null group by gaolu,case when fesi<=1.25 then '炼钢铁' else '铸造铁' end";
            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    if (!dr.IsDBNull(1))
                        this.SetValue(string.Format("本日.高炉{0}.{1}", gaolu, dr.GetString(1)), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                }
            }
            dr.Close();

            //大中修
            cmd.CommandText = "select 高炉,sum(间隔) from 休风 where  trunc(时间)=:rq and (分类='大修' or 分类='中修')  and  间隔>0　and 间隔 is not null group by 高炉";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                int gaolu = dr.IsDBNull(0) ? 6 : dr.GetInt32(0);
                this.SetValue(string.Format("本日.高炉{0}.大中修", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));

            }
            dr.Close();

            //合格铁
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where 合格铁(fesi,fes)=1 and trunc(zdsj)=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("本日.高炉{0}.合格铁", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();

            //送炼钢厂,本厂铸造,SI小,SI大
            cmd.CommandText = "select gaolu,decode(quchu,'炼钢','送炼钢厂','本厂铸造'),case when fesi<=1.25 then 'SI小' else 'SI大' end,sum(feliang) from ddluci where trunc(zdsj)=:rq and dksj is not null group by gaolu,decode(quchu,'炼钢','送炼钢厂','本厂铸造'),case when fesi<=1.25 then 'SI小' else 'SI大' end";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("本日.高炉{0}.{1}{2}", gaolu, dr.GetString(1), dr.GetString(2)), dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                }
            }
            dr.Close();

            ////炼钢铁、铸造铁成分
            cmd.CommandText = "select nvl(gaolu,9999),case when fesi<=1.25 then '炼钢铁' else '铸造铁' end,round(sum(feliang*FESI)/sum(feliang),2),round(sum(feliang*FEMn)/sum(feliang),2),round(sum(feliang*FEP)/sum(feliang),3),round(sum(feliang*FES)/sum(feliang),3) from ddluci where feliang>0 and trunc(zdsj)=:rq and dksj is not null group by rollup(gaolu),case when fesi<=1.25 then '炼钢铁' else '铸造铁' end";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    if (gaolu != 9999)
                    {
                        this.SetValue(string.Format("本日.高炉{0}.{1}{2}", gaolu, dr.GetString(1), "Si"), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                        this.SetValue(string.Format("本日.高炉{0}.{1}{2}", gaolu, dr.GetString(1), "Mn"), dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                        this.SetValue(string.Format("本日.高炉{0}.{1}{2}", gaolu, dr.GetString(1), "P"), dr.IsDBNull(4) ? 0 : dr.GetDouble(4));
                        this.SetValue(string.Format("本日.高炉{0}.{1}{2}", gaolu, dr.GetString(1), "S"), dr.IsDBNull(5) ? 0 : dr.GetDouble(5));
                    }
                    else
                    {
                        this.SetValue(string.Format("本日.{0}{1}", dr.GetString(1), "Si"), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                        this.SetValue(string.Format("本日.{0}{1}", dr.GetString(1), "Mn"), dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                        this.SetValue(string.Format("本日.{0}{1}", dr.GetString(1), "P"), dr.IsDBNull(4) ? 0 : dr.GetDouble(4));
                        this.SetValue(string.Format("本日.{0}{1}", dr.GetString(1), "S"), dr.IsDBNull(5) ? 0 : dr.GetDouble(5));
                    }
                }
            }
            dr.Close();

            //炉渣成分
            cmd.CommandText = "select  nvl(gaolu,9999),trunc(avg(ZHAR2),2) from ddluci where trunc(zdsj)=:rq and dksj is not null group by rollup(gaolu)";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    if (gaolu != 9999)
                    {
                        this.SetValue(string.Format("本日.高炉{0}.炉渣碱度", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                    }
                    else
                    {
                        this.SetValue(string.Format("本日.炉渣碱度"), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                    }
                }
            }
            dr.Close();

            //全铁产量
      
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where  trunc(zdsj)=:rq and dksj is not null  group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                   
                    this.SetValue(string.Format("本日.高炉{0}.全铁产量", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();
            //s<0.02比例
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where  trunc(zdsj)=:rq and dksj is not null and fes<=0.02 group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("本日.高炉{0}.S小于002", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();
            //P<0.009比例
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where  trunc(zdsj)=:rq and dksj is not null and fep<=0.09 group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    double fesliang = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    this.SetValue(string.Format("本日.高炉{0}.P小于009", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();
            //Ti<0.055比例
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where  trunc(zdsj)=:rq and dksj is not null and feti<=0.055 group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    double fesliang = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    this.SetValue(string.Format("本日.高炉{0}.TI小于055", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();
            //折算产量

            cmd.CommandText = "select gaolu,sum(折算产量(feliang,fesi,fes)) from ddluci where trunc(zdsj)=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("本日.高炉{0}.折算产量", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();

            cmd.CommandText = "select P05二号皮带,P06三号皮带,P07总返矿,自产湿焦,落地湿焦,P02自产焦水分,P03落地焦水分,P04焦粉水分 from 全厂日消耗 where P01日期=:rq";
            dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                this.SetValue(string.Format("本日.二号皮带"), dr.IsDBNull(0) ? 0 : dr.GetDouble(0));
                this.SetValue(string.Format("本日.三号皮带"), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                this.SetValue(string.Format("本日.返矿"), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                this.SetValue(string.Format("本日.自产湿焦"), dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                this.SetValue(string.Format("本日.落地湿焦"), dr.IsDBNull(4) ? 0 : dr.GetDouble(4));
                this.SetValue(string.Format("本日.自产湿焦水份"), dr.IsDBNull(5) ? 0 : dr.GetDouble(5));
                this.SetValue(string.Format("本日.落地湿焦水份"), dr.IsDBNull(6) ? 0 : dr.GetDouble(6));
                this.SetValue(string.Format("本日.湿焦粉水份"), dr.IsDBNull(7) ? 0 : dr.GetDouble(7));
              //  this.SetValue(string.Format("本日.湿焦粉水份"), dr.IsDBNull(8) ? 0 : dr.GetDouble(8));
            }
            dr.Close();

            // 原料消耗
            cmd.CommandText = "select 高炉,trunc(机烧矿,2),trunc(球团矿,2),trunc(国内球团矿,2),trunc(进口球团矿,2),trunc(PB块,2),trunc(纽曼块,2),trunc(其它块矿,2),trunc(高钛球团矿,2),trunc(高品位钛球,2),trunc(钛球,2),trunc(锰矿,2)," +
               "工艺称,煤粉,焦丁,trunc(自产湿焦,2),trunc(落地湿焦,2),trunc(硅石,2),trunc(萤石,2),trunc(蛇纹石,2),trunc(其它熔剂,2) from 原料消耗 where 日期=:rq";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("本日.高炉{0}.机烧", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                    this.SetValue(string.Format("本日.高炉{0}.球团矿", gaolu), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                    this.SetValue(string.Format("本日.高炉{0}.国内球团矿", gaolu), dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                    this.SetValue(string.Format("本日.高炉{0}.进口球团矿", gaolu), dr.IsDBNull(4) ? 0 : dr.GetDouble(4));
                    this.SetValue(string.Format("本日.高炉{0}.PB块", gaolu), dr.IsDBNull(5) ? 0 : dr.GetDouble(5));
                    this.SetValue(string.Format("本日.高炉{0}.纽曼块", gaolu), dr.IsDBNull(6) ? 0 : dr.GetDouble(6));
                    this.SetValue(string.Format("本日.高炉{0}.其它块矿", gaolu), dr.IsDBNull(7) ? 0 : dr.GetDouble(7));
                    this.SetValue(string.Format("本日.高炉{0}.高钛球团矿", gaolu), dr.IsDBNull(8) ? 0 : dr.GetDouble(8));
                    this.SetValue(string.Format("本日.高炉{0}.高品位钛球", gaolu), dr.IsDBNull(9) ? 0 : dr.GetDouble(9));
                    this.SetValue(string.Format("本日.高炉{0}.钛球", gaolu), dr.IsDBNull(10) ? 0 : dr.GetDouble(10));
                    this.SetValue(string.Format("本日.高炉{0}.锰矿", gaolu), dr.IsDBNull(11) ? 0 : dr.GetDouble(11));
                    this.SetValue(string.Format("本日.高炉{0}.工艺称焦炭", gaolu), dr.IsDBNull(12) ? 0 : dr.GetDouble(12));
                    this.SetValue(string.Format("本日.高炉{0}.煤粉总耗", gaolu), dr.IsDBNull(13) ? 0 : dr.GetDouble(13));
                    this.SetValue(string.Format("本日.高炉{0}.焦丁", gaolu), dr.IsDBNull(14) ? 0 : dr.GetDouble(14));
                    this.SetValue(string.Format("本日.高炉{0}.自产湿焦", gaolu), dr.IsDBNull(15) ? 0 : dr.GetDouble(15));
                    this.SetValue(string.Format("本日.高炉{0}.落地湿焦", gaolu), dr.IsDBNull(16) ? 0 : dr.GetDouble(16));
                    this.SetValue(string.Format("本日.高炉{0}.硅石", gaolu), dr.IsDBNull(17) ? 0 : dr.GetDouble(17));

                    this.SetValue(string.Format("本日.高炉{0}.萤石", gaolu), dr.IsDBNull(18) ? 0 : dr.GetDouble(18));
                    this.SetValue(string.Format("本日.高炉{0}.蛇纹石", gaolu), dr.IsDBNull(19) ? 0 : dr.GetDouble(19));
                    this.SetValue(string.Format("本日.高炉{0}.其它熔剂", gaolu), dr.IsDBNull(20) ? 0 : dr.GetDouble(20));
                   
                   
                }
            }
            dr.Close();

            // cmd.CommandText = "select gaolu,shijiaofen,fuyang,lengfengliuliang,feitie,fengwen,ludingwendu,refengyali,ludingyali,fengkoudatao,fengkouzhongtao,fengkouxiaotao,zhakoudatao,zhakouzhongtao,zhakouxiaotao,zuoliao,xuanliao,bengliao,trunc(shijiaofen*(1-全厂日消耗.P04焦粉水分/100),2) as 干焦粉  from xiaohao,全厂日消耗 where 全厂日消耗.P01日期=xiaohao.RQ and  xiaohao.RQ=:rq";
            cmd.CommandText = "select gaolu,shijiaofen,fuyang,lengfengliuliang,feitie,fengwen,ludingwendu,refengyali,ludingyali,fengkoudatao,fengkouzhongtao,fengkouxiaotao,zhakoudatao,zhakouzhongtao,zhakouxiaotao,zuoliao,xuanliao,bengliao,jiaoding,Jiaodingshuifen,JiaofenShuiFen  from xiaohao where RQ=:rq";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("本日.高炉{0}.湿焦粉", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                    this.SetValue(string.Format("本日.高炉{0}.富氧量", gaolu), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                    this.SetValue(string.Format("本日.高炉{0}.冷风流量", gaolu), dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                    this.SetValue(string.Format("本日.高炉{0}.废铁总耗", gaolu), dr.IsDBNull(4) ? 0 : dr.GetDouble(4));
                    this.SetValue(string.Format("本日.高炉{0}.平均风温", gaolu), dr.IsDBNull(5) ? 0 : dr.GetDouble(5));
                    this.SetValue(string.Format("本日.高炉{0}.炉顶温度", gaolu), dr.IsDBNull(6) ? 0 : dr.GetDouble(6));
                    this.SetValue(string.Format("本日.高炉{0}.热风压力", gaolu), dr.IsDBNull(7) ? 0 : dr.GetDouble(7));
                    this.SetValue(string.Format("本日.高炉{0}.炉顶压力", gaolu), dr.IsDBNull(8) ? 0 : dr.GetDouble(8));
                    this.SetValue(string.Format("本日.高炉{0}.风口损坏数大", gaolu), dr.IsDBNull(9) ? 0 : dr.GetDouble(9));
                    this.SetValue(string.Format("本日.高炉{0}.风口损坏数中", gaolu), dr.IsDBNull(10) ? 0 : dr.GetDouble(10));
                    this.SetValue(string.Format("本日.高炉{0}.风口损坏数小", gaolu), dr.IsDBNull(11) ? 0 : dr.GetDouble(11));
                    this.SetValue(string.Format("本日.高炉{0}.渣口损坏数大", gaolu), dr.IsDBNull(12) ? 0 : dr.GetDouble(12));
                    this.SetValue(string.Format("本日.高炉{0}.渣口损坏数中", gaolu), dr.IsDBNull(13) ? 0 : dr.GetDouble(13));
                    this.SetValue(string.Format("本日.高炉{0}.渣口损坏数小", gaolu), dr.IsDBNull(14) ? 0 : dr.GetDouble(14));
                    this.SetValue(string.Format("本日.高炉{0}.坐料次数", gaolu), dr.IsDBNull(15) ? 0 : dr.GetDouble(15));
                    this.SetValue(string.Format("本日.高炉{0}.悬料次数", gaolu), dr.IsDBNull(16) ? 0 : dr.GetDouble(16));
                    this.SetValue(string.Format("本日.高炉{0}.崩料次数", gaolu), dr.IsDBNull(17) ? 0 : dr.GetDouble(17));
                    this.SetValue(string.Format("本日.高炉{0}.湿焦丁", gaolu), dr.IsDBNull(18) ? 0 : dr.GetDouble(18));
                    this.SetValue(string.Format("本日.高炉{0}.焦丁水份", gaolu), dr.IsDBNull(19) ? 0 : dr.GetDouble(19));
                    this.SetValue(string.Format("本日.高炉{0}.焦粉水份", gaolu), dr.IsDBNull(20) ? 0 : dr.GetDouble(20));
                }
            }
            dr.Close();

            cmd.CommandText = "select 高炉,sum(间隔) from 休风 where trunc(时间)=:rq group by 高炉";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("本日.高炉{0}.休风情况", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();

            cmd.CommandText = "select 高炉,sum(间隔) from 慢风 where trunc(时间)=:rq group by 高炉";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("本日.高炉{0}.慢风", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();
            //纽曼块
            //cmd.CommandText = "select gaolu,nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0) from rbxiaohao where trunc(sj)=:rq and (MC like '%纽曼%' or beizhu like '%纽曼%') ";
            //dr = cmd.ExecuteReader();
            //while (dr.Read())
            //{
            //    if (!dr.IsDBNull(0))
            //    {
            //        int gaolu = dr.GetInt32(0);
            //        this.SetValue(string.Format("本日.高炉{0}.纽曼块", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
            //    }
            //}
            //dr.Close();
            ////PB块
            //cmd.CommandText = "select gaolu,nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0) from rbxiaohao where trunc(sj)=:rq and (MC like '%PB%' or beizhu like '%PB%') ";
            //dr = cmd.ExecuteReader();
            //while (dr.Read())
            //{
            //    if (!dr.IsDBNull(0))
            //    {
            //        int gaolu = dr.GetInt32(0);
            //        this.SetValue(string.Format("本日.高炉{0}.PB块", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
            //    }
            //}
            //dr.Close();
            //钛球
            //cmd.CommandText = "select gaolu,nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0) from rbxiaohao where trunc(sj)=:rq and (MC like '%钛球%' or beizhu like '%钛球%') ";
            //dr = cmd.ExecuteReader();
            //while (dr.Read())
            //{
            //    if (!dr.IsDBNull(0))
            //    {
            //        int gaolu = dr.GetInt32(0);
            //        this.SetValue(string.Format("本日.高炉{0}.钛球", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
            //    }
            //}
            //dr.Close();
         
            ////锰矿
            //cmd.CommandText = "select gaolu,nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0) from rbxiaohao where trunc(sj)=:rq  and (MC like '%锰矿%' or beizhu like '%锰矿%') ";
            //dr = cmd.ExecuteReader();
            //while (dr.Read())
            //{
            //    if (!dr.IsDBNull(0))
            //    {
            //        int gaolu = dr.GetInt32(0);
            //        this.SetValue(string.Format("本日.高炉{0}.锰矿", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
            //    }
            //}
            //dr.Close();
            //硅石
           // cmd.CommandText = "select gaolu,nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0) from rbxiaohao where trunc(sj)=:rq  and (MC like '%硅石%' or beizhu like '%硅石%') ";
            //dr = cmd.ExecuteReader();
            //while (dr.Read())
            //{
            //    if (!dr.IsDBNull(0))
            //    {
            //        int gaolu = dr.GetInt32(0);
            //        this.SetValue(string.Format("本日.高炉{0}.硅石", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
            //    }
            //}
            //dr.Close();
            ////白云石
            //cmd.CommandText = "select gaolu,nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0) from rbxiaohao where trunc(sj)=:rq  and (MC like '%白云石%' or beizhu like '%白云石%') ";
            //dr = cmd.ExecuteReader();
            //while (dr.Read())
            //{
            //    if (!dr.IsDBNull(0))
            //    {
            //        int gaolu = dr.GetInt32(0);
            //        this.SetValue(string.Format("本日.高炉{0}.白云石", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
            //    }
            //}
            //dr.Close();
            ////蛇纹石
            //cmd.CommandText = "select gaolu,nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0) from rbxiaohao where trunc(sj)=:rq  and (MC like '%蛇纹石%' or beizhu like '%蛇纹石%') ";
            //dr = cmd.ExecuteReader();
            //while (dr.Read())
            //{
            //    if (!dr.IsDBNull(0))
            //    {
            //        int gaolu = dr.GetInt32(0);
            //        this.SetValue(string.Format("本日.高炉{0}.蛇纹石", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
            //    }
            //}
            //dr.Close();
            #endregion

            #region 累计部分

            //炼钢铁
            cmd.CommandText = "select gaolu,case when fesi<=1.25 then '炼钢铁' else '铸造铁' end,sum(feliang) from ddluci where trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and dksj is not null group by gaolu,case when fesi<=1.25 then '炼钢铁' else '铸造铁' end";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    if (!dr.IsDBNull(1))
                        this.SetValue(string.Format("累计.高炉{0}.{1}", gaolu, dr.GetString(1)), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                }
            }
            dr.Close();

            //大中修
            cmd.CommandText = "select 高炉,sum(间隔) from 休风 where   trunc(时间,'MONTH')=trunc(:rq,'MONTH') and trunc(时间)<=:rq and (分类='大修' or 分类='中修')  and  间隔>0　and 间隔 is not null group by 高炉";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                int gaolu = dr.IsDBNull(0) ? 6 : dr.GetInt32(0);
                this.SetValue(string.Format("累计.高炉{0}.大中修", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));

            }
            dr.Close();

            //合格铁
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where 合格铁(fesi,fes)=1 and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and  trunc(zdsj)<=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("累计.高炉{0}.合格铁", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();

            //送炼钢厂,本厂铸造,SI小,SI大
            cmd.CommandText = "select gaolu,decode(quchu,'炼钢','送炼钢厂','本厂铸造'),case when fesi<=1.25 then 'SI小' else 'SI大' end,sum(feliang) from ddluci where  trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and dksj is not null group by gaolu,decode(quchu,'炼钢','送炼钢厂','本厂铸造'),case when fesi<=1.25 then 'SI小' else 'SI大' end";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("累计.高炉{0}.{1}{2}", gaolu, dr.GetString(1), dr.GetString(2)), dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                }
            }
            dr.Close();

            //炼钢铁、铸造铁成分
            cmd.CommandText = "select nvl(gaolu,9999),case when fesi<=1.25 then '炼钢铁' else '铸造铁' end,round(sum(feliang*FESI)/sum(feliang),2),round(sum(feliang*FEMn)/sum(feliang),2),round(sum(feliang*FEP)/sum(feliang),3),round(sum(feliang*FES)/sum(feliang),3) from ddluci where feliang>0 and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and dksj is not null group by rollup(gaolu),case when fesi<=1.25 then '炼钢铁' else '铸造铁' end";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    if (gaolu != 9999)
                    {
                        this.SetValue(string.Format("累计.高炉{0}.{1}{2}", gaolu, dr.GetString(1), "Si"), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                        this.SetValue(string.Format("累计.高炉{0}.{1}{2}", gaolu, dr.GetString(1), "Mn"), dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                        this.SetValue(string.Format("累计.高炉{0}.{1}{2}", gaolu, dr.GetString(1), "P"), dr.IsDBNull(4) ? 0 : dr.GetDouble(4));
                        this.SetValue(string.Format("累计.高炉{0}.{1}{2}", gaolu, dr.GetString(1), "S"), dr.IsDBNull(5) ? 0 : dr.GetDouble(5));
                    }
                    else
                    {
                        this.SetValue(string.Format("累计.{0}{1}", dr.GetString(1), "Si"), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                        this.SetValue(string.Format("累计.{0}{1}", dr.GetString(1), "Mn"), dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                        this.SetValue(string.Format("累计.{0}{1}", dr.GetString(1), "P"), dr.IsDBNull(4) ? 0 : dr.GetDouble(4));
                        this.SetValue(string.Format("累计.{0}{1}", dr.GetString(1), "S"), dr.IsDBNull(5) ? 0 : dr.GetDouble(5));
                    }
                }
            }
            dr.Close();

            //炉渣成分
            cmd.CommandText = "select nvl(gaolu,9999),trunc(avg(ZHAR2),2) from ddluci where  trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and dksj is not null group by rollup(gaolu)";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    if (gaolu != 9999)
                    {
                        this.SetValue(string.Format("累计.高炉{0}.炉渣碱度", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                    }
                    else
                    {
                        this.SetValue(string.Format("累计.炉渣碱度"), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                    }
                }
            }
            dr.Close();

            //全铁产量
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where  trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("累计.高炉{0}.全铁产量", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();

            //s<0.02
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where  trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and fes<=0.02 and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("累计.高炉{0}.S小于002", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();
            //p<0.002
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where  trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and fep<=0.09 and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("累计.高炉{0}.P小于009", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();
            //p<0.0055
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where  trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and feti<=0.055 and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("累计.高炉{0}.TI小于055", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();
            //折算产量

            cmd.CommandText = "select gaolu,sum(折算产量(feliang,fesi,fes)) from ddluci where trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and  trunc(zdsj)<=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("累计.高炉{0}.折算产量", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();

            cmd.CommandText = "select SUM(P05二号皮带),SUM(P06三号皮带),SUM(P07总返矿),SUM(自产湿焦),SUM(落地湿焦),TRUNC(AVG(P02自产焦水分),2),TRUNC(AVG(P03落地焦水分),2),TRUNC(AVG(P04焦粉水分),2) from 全厂日消耗 where trunc(P01日期,'MONTH')=trunc(:rq,'MONTH') and P01日期<=:rq";
            dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                this.SetValue(string.Format("累计.二号皮带"), dr.IsDBNull(0) ? 0 : dr.GetDouble(0));
                this.SetValue(string.Format("累计.三号皮带"), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                this.SetValue(string.Format("累计.返矿"), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                this.SetValue(string.Format("累计.自产湿焦"), dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                this.SetValue(string.Format("累计.落地湿焦"), dr.IsDBNull(4) ? 0 : dr.GetDouble(4));
                this.SetValue(string.Format("累计.自产湿焦水份"), dr.IsDBNull(5) ? 0 : dr.GetDouble(5));
                this.SetValue(string.Format("累计.落地湿焦水份"), dr.IsDBNull(6) ? 0 : dr.GetDouble(6));
                this.SetValue(string.Format("累计.湿焦粉水份"), dr.IsDBNull(7) ? 0 : dr.GetDouble(7));
            }
            dr.Close();

            cmd.CommandText = "select 高炉,sum(trunc(机烧矿,2)),sum(trunc(球团矿,2)),sum(trunc(国内球团矿,2)),sum(trunc(进口球团矿,2)),sum(trunc(PB块,2)),sum(trunc(纽曼块,2)),sum(trunc(其它块矿,2)),sum(trunc(高钛球团矿,2)),sum(trunc(高品位钛球,2)),sum(trunc(钛球,2)),sum(trunc(锰矿,2)),"
                + "sum(工艺称),sum(煤粉),sum(焦丁),sum(trunc(自产湿焦,2)),sum(trunc(落地湿焦,2)),sum(trunc(硅石,2)),sum(trunc(萤石,2)),sum(trunc(蛇纹石,2)),sum(trunc(其它熔剂,2))"
                +" from 原料消耗 where  trunc(日期,'MONTH')=trunc(:rq,'MONTH') and 日期<=:rq group by 高炉";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("累计.高炉{0}.机烧", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                    this.SetValue(string.Format("累计.高炉{0}.球团矿", gaolu), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                    this.SetValue(string.Format("累计.高炉{0}.国内球团矿", gaolu), dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                    this.SetValue(string.Format("累计.高炉{0}.进口球团矿", gaolu), dr.IsDBNull(4) ? 0 : dr.GetDouble(4));
                    this.SetValue(string.Format("累计.高炉{0}.PB块", gaolu), dr.IsDBNull(5) ? 0 : dr.GetDouble(5));
                    this.SetValue(string.Format("累计.高炉{0}.纽曼块", gaolu), dr.IsDBNull(6) ? 0 : dr.GetDouble(6));
                    this.SetValue(string.Format("累计.高炉{0}.其它块矿", gaolu), dr.IsDBNull(7) ? 0 : dr.GetDouble(7));
                    this.SetValue(string.Format("累计.高炉{0}.高钛球团矿", gaolu), dr.IsDBNull(8) ? 0 : dr.GetDouble(8));
                    this.SetValue(string.Format("累计.高炉{0}.高品位钛球", gaolu), dr.IsDBNull(9) ? 0 : dr.GetDouble(9));
                    this.SetValue(string.Format("累计.高炉{0}.钛球", gaolu), dr.IsDBNull(10) ? 0 : dr.GetDouble(10));
                    this.SetValue(string.Format("累计.高炉{0}.锰矿", gaolu), dr.IsDBNull(11) ? 0 : dr.GetDouble(11));
                    this.SetValue(string.Format("累计.高炉{0}.工艺称焦炭", gaolu), dr.IsDBNull(12) ? 0 : dr.GetDouble(12));
                    this.SetValue(string.Format("累计.高炉{0}.煤粉总耗", gaolu), dr.IsDBNull(13) ? 0 : dr.GetDouble(13));
                    this.SetValue(string.Format("累计.高炉{0}.焦丁", gaolu), dr.IsDBNull(14) ? 0 : dr.GetDouble(14));
                    this.SetValue(string.Format("累计.高炉{0}.自产湿焦", gaolu), dr.IsDBNull(15) ? 0 : dr.GetDouble(15));
                    this.SetValue(string.Format("累计.高炉{0}.落地湿焦", gaolu), dr.IsDBNull(16) ? 0 : dr.GetDouble(16));
                    this.SetValue(string.Format("累计.高炉{0}.硅石", gaolu), dr.IsDBNull(17) ? 0 : dr.GetDouble(17));

                    this.SetValue(string.Format("累计.高炉{0}.萤石", gaolu), dr.IsDBNull(18) ? 0 : dr.GetDouble(18));
                    this.SetValue(string.Format("累计.高炉{0}.蛇纹石", gaolu), dr.IsDBNull(19) ? 0 : dr.GetDouble(19));
                    this.SetValue(string.Format("累计.高炉{0}.其它熔剂", gaolu), dr.IsDBNull(20) ? 0 : dr.GetDouble(20));
                   
                }
            }
            dr.Close();





            // cmd.CommandText = "select gaolu,shijiaofen,fuyang,lengfengliuliang,feitie,fengwen,ludingwendu,refengyali,ludingyali,fengkoudatao,fengkouzhongtao,fengkouxiaotao,zhakoudatao,zhakouzhongtao,zhakouxiaotao,zuoliao,xuanliao,bengliao,trunc(shijiaofen*(1-全厂日消耗.P04焦粉水分/100),2) as 干焦粉  from xiaohao,全厂日消耗 where 全厂日消耗.P01日期=xiaohao.RQ and  xiaohao.RQ=:rq";
            cmd.CommandText = "select gaolu,sum(shijiaofen),sum(fuyang),sum(lengfengliuliang),sum(feitie),trunc(avg(fengwen),0),trunc(avg(ludingwendu),0),trunc(avg(refengyali),0),trunc(avg(ludingyali),0),sum(fengkoudatao),sum(fengkouzhongtao),sum(fengkouxiaotao),sum(zhakoudatao),sum(zhakouzhongtao),sum(zhakouxiaotao),sum(zuoliao),sum(xuanliao),sum(bengliao), round(sum(jiaoding*(100-JiaodingShuiFen)/100),2), round(sum(shijiaofen*(100-JiaoFenShuiFen)/100),2)  from xiaohao where trunc(RQ,'MONTH')=trunc(:rq,'MONTH') and RQ<=:rq  group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("累计.高炉{0}.湿焦粉", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                    this.SetValue(string.Format("累计.高炉{0}.富氧量", gaolu), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                    this.SetValue(string.Format("累计.高炉{0}.冷风流量", gaolu), dr.IsDBNull(3) ? 0 : dr.GetDouble(3));
                    this.SetValue(string.Format("累计.高炉{0}.废铁总耗", gaolu), dr.IsDBNull(4) ? 0 : dr.GetDouble(4));
                    this.SetValue(string.Format("累计.高炉{0}.平均风温", gaolu), dr.IsDBNull(5) ? 0 : dr.GetDouble(5));
                    this.SetValue(string.Format("累计.高炉{0}.炉顶温度", gaolu), dr.IsDBNull(6) ? 0 : dr.GetDouble(6));
                    this.SetValue(string.Format("累计.高炉{0}.热风压力", gaolu), dr.IsDBNull(7) ? 0 : dr.GetDouble(7));
                    this.SetValue(string.Format("累计.高炉{0}.炉顶压力", gaolu), dr.IsDBNull(8) ? 0 : dr.GetDouble(8));
                    this.SetValue(string.Format("累计.高炉{0}.风口损坏数大", gaolu), dr.IsDBNull(9) ? 0 : dr.GetDouble(9));
                    this.SetValue(string.Format("累计.高炉{0}.风口损坏数中", gaolu), dr.IsDBNull(10) ? 0 : dr.GetDouble(10));
                    this.SetValue(string.Format("累计.高炉{0}.风口损坏数小", gaolu), dr.IsDBNull(11) ? 0 : dr.GetDouble(11));
                    this.SetValue(string.Format("累计.高炉{0}.渣口损坏数大", gaolu), dr.IsDBNull(12) ? 0 : dr.GetDouble(12));
                    this.SetValue(string.Format("累计.高炉{0}.渣口损坏数中", gaolu), dr.IsDBNull(13) ? 0 : dr.GetDouble(13));
                    this.SetValue(string.Format("累计.高炉{0}.渣口损坏数小", gaolu), dr.IsDBNull(14) ? 0 : dr.GetDouble(14));
                    this.SetValue(string.Format("累计.高炉{0}.坐料次数", gaolu), dr.IsDBNull(15) ? 0 : dr.GetDouble(15));
                    this.SetValue(string.Format("累计.高炉{0}.悬料次数", gaolu), dr.IsDBNull(16) ? 0 : dr.GetDouble(16));
                    this.SetValue(string.Format("累计.高炉{0}.崩料次数", gaolu), dr.IsDBNull(17) ? 0 : dr.GetDouble(17));
                    this.SetValue(string.Format("累计.高炉{0}.干焦丁", gaolu), dr.IsDBNull(18) ? 0 : dr.GetDouble(18));
                    this.SetValue(string.Format("累计.高炉{0}.干焦粉", gaolu), dr.IsDBNull(19) ? 0 : dr.GetDouble(19));
                }
            }
            dr.Close();

            cmd.CommandText = "select 高炉,sum(间隔) from 休风 where trunc(时间,'MONTH')=trunc(:rq,'MONTH') and trunc(时间)<=:rq group by 高炉";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("累计.高炉{0}.休风情况", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();
            cmd.CommandText = "select 高炉,sum(间隔) from 慢风 where  trunc(时间,'MONTH')=trunc(:rq,'MONTH') and trunc(时间)<=:rq group by 高炉";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    this.SetValue(string.Format("累计.高炉{0}.慢风", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                }
            }
            dr.Close();
            ////纽曼块
            //cmd.CommandText = "select gaolu,sum(nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0)) from rbxiaohao where trunc(sj,'MONTH')=trunc(:rq,'MONTH') and trunc(sj)<=:rq and (MC like '%纽曼%' or beizhu like '%纽曼%') group by gaolu";
            //dr = cmd.ExecuteReader();
            //while (dr.Read())
            //{
            //    if (!dr.IsDBNull(0))
            //    {
            //        int gaolu = dr.GetInt32(0);
            //        this.SetValue(string.Format("累计.高炉{0}.纽曼块", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
            //    }
            //}
            //dr.Close();
            //FB块
            //cmd.CommandText = "select gaolu,sum(nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0)) from rbxiaohao where trunc(sj,'MONTH')=trunc(:rq,'MONTH') and trunc(sj)<=:rq and (MC like '%PB%' or beizhu like '%PB%') group by gaolu";
            //dr = cmd.ExecuteReader();
            //while (dr.Read())
            //{
            //    if (!dr.IsDBNull(0))
            //    {
            //        int gaolu = dr.GetInt32(0);
            //        this.SetValue(string.Format("累计.高炉{0}.PB块", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
            //    }
            //}
            //dr.Close();
            ////钛球
            //cmd.CommandText = "select gaolu,sum(nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0)) from rbxiaohao where trunc(sj,'MONTH')=trunc(:rq,'MONTH') and trunc(sj)<=:rq and (MC like '%钛球%' or beizhu like '%钛球%') group by gaolu";
            //dr = cmd.ExecuteReader();
            //while (dr.Read())
            //{
            //    if (!dr.IsDBNull(0))
            //    {
            //        int gaolu = dr.GetInt32(0);
            //        this.SetValue(string.Format("累计.高炉{0}.钛球", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
            //    }
            //}
            //dr.Close();

            ////锰矿
            //cmd.CommandText = "select gaolu,sum(nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0)) from rbxiaohao where trunc(sj,'MONTH')=trunc(:rq,'MONTH') and trunc(sj)<=:rq  and (MC like '%锰矿%' or beizhu like '%锰矿%') group by gaolu ";
            //dr = cmd.ExecuteReader();
            //while (dr.Read())
            //{
            //    if (!dr.IsDBNull(0))
            //    {
            //        int gaolu = dr.GetInt32(0);
            //        this.SetValue(string.Format("累计.高炉{0}.锰矿", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
            //    }
            //}
            //dr.Close();
            ////硅石
            //cmd.CommandText = "select gaolu,sum(nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0)) from rbxiaohao where trunc(sj,'MONTH')=trunc(:rq,'MONTH') and trunc(sj)<=:rq and (MC like '%硅石%' or beizhu like '%硅石%') group by gaolu";
            //dr = cmd.ExecuteReader();
            //while (dr.Read())
            //{
            //    if (!dr.IsDBNull(0))
            //    {
            //        int gaolu = dr.GetInt32(0);
            //        this.SetValue(string.Format("累计.高炉{0}.硅石", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
            //    }
            //}
            //dr.Close();
            ////白云石
            //cmd.CommandText = "select gaolu,sum(nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0)) from rbxiaohao where trunc(sj,'MONTH')=trunc(:rq,'MONTH') and trunc(sj)<=:rq  and (MC like '%白云石%' or beizhu like '%白云石%') group by gaolu ";
            //dr = cmd.ExecuteReader();
            //while (dr.Read())
            //{
            //    if (!dr.IsDBNull(0))
            //    {
            //        int gaolu = dr.GetInt32(0);
            //        this.SetValue(string.Format("累计.高炉{0}.白云石", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
            //    }
            //}
            //dr.Close();
            ////蛇纹石
            //cmd.CommandText = "select gaolu,sum(nvl(baiban,0)+nvl(zhongban,0)+nvl(yeban,0)) from rbxiaohao where trunc(sj,'MONTH')=trunc(:rq,'MONTH') and trunc(sj)<=:rq and (MC like '%蛇纹石%' or beizhu like '%蛇纹石%') group by gaolu ";
            //dr = cmd.ExecuteReader();
            //while (dr.Read())
            //{
            //    if (!dr.IsDBNull(0))
            //    {
            //        int gaolu = dr.GetInt32(0);
            //        this.SetValue(string.Format("累计.高炉{0}.蛇纹石", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
            //    }
            //}
            //dr.Close();

            if (rq.Day > 1)
            {
                cmd.CommandText = "select P01单位,P38干毛焦,P27入炉焦炭总耗 from 技术日报 where P日期+1=:rq and P02项目='累计'";
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                  
                    string gaolustr = dr.GetString(0);
                    if (gaolustr.Contains("#") && gaolustr.Length>1)
                    {    
                        int gaolu;
                        if (int.TryParse(gaolustr.Substring(0,gaolustr.Length - 1), out gaolu))
                        {
                            this.SetValue(string.Format("累计.高炉{0}.昨日干毛焦", gaolu), dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                            this.SetValue(string.Format("累计.高炉{0}.昨日入炉焦炭", gaolu), dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                        }
                    }
                    else if (gaolustr == "Q")
                    {
                        this.SetValue("累计.昨日干毛焦", dr.IsDBNull(1) ? 0 : dr.GetDouble(1));
                        this.SetValue("累计.昨日入炉焦炭", dr.IsDBNull(2) ? 0 : dr.GetDouble(2));
                    }
                }
            }

            this.SetValue("累计.累计天数", rq.Day);

            #endregion
            
            conn.Close();
           
        }
    }
}
