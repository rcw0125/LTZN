using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;

namespace LTZN.技术日报
{
    public class 原料消耗
    {
        public double[] 总耗 = new double[6];
        public double[] 机烧 = new double[6];
        public double[] 竖球 = new double[6];
        public double[] 生料 = new double[6];
        public double[] 熟料 = new double[6];
        public double[] 本溪矿 = new double[6];

        public double[] 自产湿焦 = new double[6];
        public double[] 落地湿焦 = new double[6];

        public double 自产湿焦水分 = 0;
        public double 落地湿焦水分 = 0;

        public double[] 干毛焦 = new double[6];
        public double[] 湿焦粉 = new double[6];
        public double 湿焦粉水分 = 0;

        public double[] 入炉焦炭 = new double[6];
        public double[] 煤粉 = new double[6];  //喷煤
        public double[] 焦丁 = new double[6];
        public double[] 熟料比 = new double[6];
        public double[] 综合焦炭 = new double[6];
        public double[] 综合折算焦比 = new double[6];
        public double[] 焦炭负荷 = new double[6];
        public double[] 冶炼强度 = new double[6];
        public double[] 富氧率 = new double[6];
        public double[] 废铁 = new double[6];
        public double[] 平均风温 = new double[6];
        public double[] 炉顶温度 = new double[6];
        public double[] 热风压力 = new double[6];
        public double[] 炉顶压力 = new double[6];

        public double[] 休风情况 = new double[6];
        public double[] 慢风 = new double[6];
        public double[] 坐料次数 = new double[6];
        public double[] 悬料次数 = new double[6];
        public double[] 崩料次数 = new double[6];
        public double[] 风口损坏数大 = new double[6];
        public double[] 风口损坏数中 = new double[6];
        public double[] 风口损坏数小 = new double[6];
        public double[] 渣口损坏数大 = new double[6];
        public double[] 渣口损坏数中 = new double[6];
        public double[] 渣口损坏数小 = new double[6];
        //public double[] 干焦粉 = new double[6];

        private double[] 富氧量 = new double[6]; //24小时
        private double[] 冷风流量 = new double[6]; // m3/min
        public double[] 工艺称焦炭 = new double[6];
        private double 全厂机烧 = 0;
        private double 二号皮带 = 0;
        private double 三号皮带 = 0;
        private double 返矿 = 0;

        //累计
        public double[] 累计总耗 = new double[6];
        public double[] 累计机烧 = new double[6];
        public double[] 累计竖球 = new double[6];
        public double[] 累计生料 = new double[6];
        public double[] 累计熟料 = new double[6];
        public double[] 累计本溪矿 = new double[6];

        public double[] 累计自产湿焦 = new double[6];
        public double[] 累计落地湿焦 = new double[6];

        public double 累计自产湿焦水分 = 0;
        public double 累计落地湿焦水分 = 0;

        public double[] 累计干毛焦 = new double[6];

        public double[] 累计湿焦粉 = new double[6];
        public double 累计湿焦粉水分 = 0;

        public double[] 累计入炉焦炭 = new double[6];
        public double[] 累计煤粉 = new double[6];  //喷煤
        public double[] 累计焦丁 = new double[6];
        public double[] 累计熟料比 = new double[6];
        public double[] 累计综合焦炭 = new double[6];
        public double[] 累计综合折算焦比 = new double[6];
        public double[] 累计焦炭负荷 = new double[6];
        public double[] 累计冶炼强度 = new double[6];
        public double[] 累计富氧率 = new double[6];
        public double[] 累计废铁 = new double[6];
        public double[] 累计平均风温 = new double[6];
        public double[] 累计炉顶温度 = new double[6];
        public double[] 累计热风压力 = new double[6];
        public double[] 累计炉顶压力 = new double[6];
        public double[] 累计休风情况 = new double[6];
        public double[] 累计慢风 = new double[6];
        public double[] 累计坐料次数 = new double[6];
        public double[] 累计悬料次数 = new double[6];
        public double[] 累计崩料次数 = new double[6];
        public double[] 累计风口损坏数大 = new double[6];
        public double[] 累计风口损坏数中 = new double[6];
        public double[] 累计风口损坏数小 = new double[6];
        public double[] 累计渣口损坏数大 = new double[6];
        public double[] 累计渣口损坏数中 = new double[6];
        public double[] 累计渣口损坏数小 = new double[6];
        public double[] 累计干焦粉 = new double[6];
        
        private double[] 累计富氧量 = new double[6]; //24小时
        private double[] 累计冷风流量 = new double[6]; // m3/min
        public double[] 累计工艺称焦炭 = new double[6];
        private double 累计全厂机烧 = 0;
        private double 累计二号皮带 = 0;
        private double 累计三号皮带 = 0;
        private double 累计返矿 = 0;

        public double[] 累计大中修 = new double[6];
        public double[] 大中修 = new double[6];
        public double[] 累计有效炉容 = new double[6];
        public double[] 有效炉容 = new double[6];

        public void getDataBy(DateTime rq)
        {
            rq = rq.Date;
            OracleConnection cn = new OracleConnection();
            cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
            cn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = cn;
            cmd.CommandText = "select P05二号皮带,P06三号皮带,P07总返矿,自产湿焦,落地湿焦,P02自产焦水分,P03落地焦水分,P04焦粉水分,TRUNC(NVL(自产湿焦,0)*(1-NVL(P02自产焦水分,0)/100)+NVL(落地湿焦,0)*(1-NVL(P03落地焦水分,0)/100),2) as 干毛焦 from 全厂日消耗 where P01日期=:rq";
            cmd.Parameters.Add(":rq", OracleType.DateTime).Value = rq;
            OracleDataReader  dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                二号皮带 = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
                三号皮带 = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                返矿 = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                //自产湿焦[5] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                //落地湿焦[5] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                自产湿焦水分 = dr.IsDBNull(5) ? 0 : dr.GetDouble(5);
                落地湿焦水分 = dr.IsDBNull(6) ? 0 : dr.GetDouble(6);
                湿焦粉水分=dr.IsDBNull(7) ? 0 : dr.GetDouble(7);
                干毛焦[5] = dr.IsDBNull(8) ? 0 : dr.GetDouble(8);
            }
            dr.Close();

            cmd.CommandText = "select 高炉,trunc(机烧矿,2),trunc(竖球,2),trunc(熟料,2),trunc(生料,2),trunc(本溪矿,2),工艺称,煤粉,焦丁,trunc(自产湿焦,2),trunc(落地湿焦,2) from 原料消耗 where 日期=:rq";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    机烧[gaolu-1] = dr.IsDBNull(1) ? 0 : Convert.ToInt32(dr.GetDouble(1));
                    竖球[gaolu - 1] = dr.IsDBNull(2) ? 0 : Convert.ToInt32(dr.GetDouble(2));
                    熟料[gaolu - 1] = dr.IsDBNull(3) ? 0 : Convert.ToInt32(dr.GetDouble(3));
                    生料[gaolu - 1] = dr.IsDBNull(4) ? 0 : Convert.ToInt32(dr.GetDouble(4));
                    本溪矿[gaolu - 1] = dr.IsDBNull(5) ? 0 : Convert.ToInt32(dr.GetDouble(5));
                    工艺称焦炭[gaolu - 1] = dr.IsDBNull(6) ? 0 : dr.GetDouble(6);
                    煤粉[gaolu - 1] = dr.IsDBNull(7) ? 0 : dr.GetDouble(7);
                    焦丁[gaolu - 1] = dr.IsDBNull(8) ? 0 : dr.GetDouble(8);
                    自产湿焦[gaolu - 1] = dr.IsDBNull(9) ? 0 : dr.GetDouble(9);
                    落地湿焦[gaolu - 1] = dr.IsDBNull(10) ? 0 : dr.GetDouble(10);
                   
                }
            }
            dr.Close();

           // cmd.CommandText = "select gaolu,shijiaofen,fuyang,lengfengliuliang,feitie,fengwen,ludingwendu,refengyali,ludingyali,fengkoudatao,fengkouzhongtao,fengkouxiaotao,zhakoudatao,zhakouzhongtao,zhakouxiaotao,zuoliao,xuanliao,bengliao,trunc(shijiaofen*(1-全厂日消耗.P04焦粉水分/100),2) as 干焦粉  from xiaohao,全厂日消耗 where 全厂日消耗.P01日期=xiaohao.RQ and  xiaohao.RQ=:rq";
            cmd.CommandText = "select gaolu,shijiaofen,fuyang,lengfengliuliang,feitie,fengwen,ludingwendu,refengyali,ludingyali,fengkoudatao,fengkouzhongtao,fengkouxiaotao,zhakoudatao,zhakouzhongtao,zhakouxiaotao,zuoliao,xuanliao,bengliao  from xiaohao where RQ=:rq";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    湿焦粉[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    富氧量[gaolu - 1] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                    冷风流量[gaolu - 1] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                    废铁[gaolu - 1] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                    平均风温[gaolu - 1] = dr.IsDBNull(5) ? 0 : dr.GetDouble(5);
                    炉顶温度[gaolu - 1] = dr.IsDBNull(6) ? 0 : dr.GetDouble(6);
                    热风压力[gaolu - 1] = dr.IsDBNull(7) ? 0 : dr.GetDouble(7);
                    炉顶压力[gaolu - 1] = dr.IsDBNull(8) ? 0 : dr.GetDouble(8);
                    风口损坏数大[gaolu - 1] = dr.IsDBNull(9) ? 0 : dr.GetDouble(9);
                    风口损坏数中[gaolu - 1] = dr.IsDBNull(10) ? 0 : dr.GetDouble(10);
                    风口损坏数小[gaolu - 1] = dr.IsDBNull(11) ? 0 : dr.GetDouble(11);
                    渣口损坏数大[gaolu - 1] = dr.IsDBNull(12) ? 0 : dr.GetDouble(12);
                    渣口损坏数中[gaolu - 1] = dr.IsDBNull(13) ? 0 : dr.GetDouble(13);
                    渣口损坏数小[gaolu - 1] = dr.IsDBNull(14) ? 0 : dr.GetDouble(14);
                    坐料次数[gaolu - 1] = dr.IsDBNull(15) ? 0 : dr.GetDouble(15);
                    悬料次数[gaolu - 1] = dr.IsDBNull(16) ? 0 : dr.GetDouble(16);
                    崩料次数[gaolu - 1] = dr.IsDBNull(17) ? 0 : dr.GetDouble(17);
                    //干焦粉[gaolu - 1]= dr.IsDBNull(18) ? 0 : dr.GetDouble(18);
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
                    休风情况[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
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
                    慢风[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                }
            }
            dr.Close();

            //大中修
            cmd.CommandText = "select 高炉,sum(间隔) from 休风 where  trunc(时间)=:rq and (分类='大修' or 分类='中修')  and  间隔>0　and 间隔 is not null group by 高炉";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                int gaolu = dr.IsDBNull(0) ? 6 : dr.GetInt32(0);
                if (gaolu < 6) 大中修[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
            }
            大中修[5] = 大中修[0] + 大中修[1] + 大中修[2] + 大中修[3] + 大中修[4];
            dr.Close();

            for (int i = 0; i < 5; i++)
            {
                休风情况[5] += 休风情况[i];
                慢风[5] += 慢风[i];
                坐料次数[5] += 坐料次数[i];
                悬料次数[5] += 悬料次数[i];
                崩料次数[5] += 崩料次数[i];
                风口损坏数大[5] += 风口损坏数大[i];
                风口损坏数中[5] += 风口损坏数中[i];
                风口损坏数小[5] += 风口损坏数小[i];
                渣口损坏数大[5] += 渣口损坏数大[i];
                渣口损坏数中[5] += 渣口损坏数中[i];
                渣口损坏数小[5] += 渣口损坏数小[i];
                废铁[5] += 废铁[i];
                机烧[5] += 机烧[i];
                竖球[5] += 竖球[i];
                熟料[5] += 熟料[i];
                生料[5] += 生料[i];
                本溪矿[5] += 本溪矿[i];
                工艺称焦炭[5] += 工艺称焦炭[i];
                湿焦粉[5] += 湿焦粉[i];
                煤粉[5] += 煤粉[i];
                焦丁[5] += 焦丁[i];
                富氧量[5] += 富氧量[i];
                冷风流量[5] += 冷风流量[i];
                热风压力[5] += 热风压力[i];
                炉顶温度[5] += 炉顶温度[i];
                炉顶压力[5] += 炉顶压力[i];
                自产湿焦[5] += 自产湿焦[i];
                落地湿焦[5] += 落地湿焦[i];
               // 干焦粉[5] += 干焦粉[i];

            }

            有效炉容[0] = (((1440 * 1) - 大中修[0]) / (1440 * 1)) * Constants.volume1;
            有效炉容[1] = (((1440 * 1) - 大中修[1]) / (1440 * 1)) * Constants.volume2;
            有效炉容[2] = (((1440 * 1) - 大中修[2]) / (1440 * 1)) * Constants.volume3;
            有效炉容[3] = (((1440 * 1) - 大中修[3]) / (1440 * 1)) * Constants.volume4;
            有效炉容[4] = (((1440 * 1) - 大中修[4]) / (1440 * 1)) * Constants.volume5;
            有效炉容[5] = 有效炉容[0] + 有效炉容[1] + 有效炉容[2] + 有效炉容[3] + 有效炉容[4];

            平均风温[5] = (平均风温[0] * Convert.ToInt32(平均风温[0] > 0) + 平均风温[1] * Convert.ToInt32(平均风温[1] > 0) + 平均风温[2] * Convert.ToInt32(平均风温[2] > 0) + 平均风温[3] * Convert.ToInt32(平均风温[3] > 0) + 平均风温[4] * Convert.ToInt32(平均风温[4] > 0)) / (Convert.ToInt32(平均风温[0] > 0) + Convert.ToInt32(平均风温[1] > 0) + Convert.ToInt32(平均风温[2] > 0) + Convert.ToInt32(平均风温[3] > 0) + Convert.ToInt32(平均风温[4] > 0));
            热风压力[5] = (热风压力[0] * Convert.ToInt32(热风压力[0] > 0) + 热风压力[1] * Convert.ToInt32(热风压力[1] > 0) + 热风压力[2] * Convert.ToInt32(热风压力[2] > 0) + 热风压力[3] * Convert.ToInt32(热风压力[3] > 0) + 热风压力[4] * Convert.ToInt32(热风压力[4] > 0)) / (Convert.ToInt32(热风压力[0] > 0) + Convert.ToInt32(热风压力[1] > 0) + Convert.ToInt32(热风压力[2] > 0) + Convert.ToInt32(热风压力[3] > 0) + Convert.ToInt32(热风压力[4] > 0));
            炉顶温度[5] = (炉顶温度[0] * Convert.ToInt32(炉顶温度[0] > 0) + 炉顶温度[1] * Convert.ToInt32(炉顶温度[1] > 0) + 炉顶温度[2] * Convert.ToInt32(炉顶温度[2] > 0) + 炉顶温度[3] * Convert.ToInt32(炉顶温度[3] > 0) + 炉顶温度[4] * Convert.ToInt32(炉顶温度[4] > 0)) / (Convert.ToInt32(炉顶温度[0] > 0) + Convert.ToInt32(炉顶温度[1] > 0) + Convert.ToInt32(炉顶温度[2] > 0) + Convert.ToInt32(炉顶温度[3] > 0) + Convert.ToInt32(炉顶温度[4] > 0));
            炉顶压力[5] = (炉顶压力[0] * Convert.ToInt32(炉顶压力[0] > 0) + 炉顶压力[1] * Convert.ToInt32(炉顶压力[1] > 0) + 炉顶压力[2] * Convert.ToInt32(炉顶压力[2] > 0) + 炉顶压力[3] * Convert.ToInt32(炉顶压力[3] > 0) + 炉顶压力[4] * Convert.ToInt32(炉顶压力[4] > 0)) / (Convert.ToInt32(炉顶压力[0] > 0) + Convert.ToInt32(炉顶压力[1] > 0) + Convert.ToInt32(炉顶压力[2] > 0) + Convert.ToInt32(炉顶压力[3] > 0) + Convert.ToInt32(炉顶压力[4] > 0));

            全厂机烧 = 二号皮带 + 三号皮带 - 返矿 + ((竖球[5] + 熟料[5] + 生料[5] + 本溪矿[5]) / 0.92 - (竖球[5] + 熟料[5] + 生料[5] + 本溪矿[5]));
            
            for (int i = 0; i < 5; i++)
            {
                if (机烧[5] > 0)
                  机烧[i] = double.Parse(((double)(全厂机烧 * 机烧[i] / 机烧[5])).ToString ("########0")); //分配机烧
              
              //if (工艺称焦炭[5] > 0)
              //{
              //    自产湿焦[i] = 自产湿焦[5] * 工艺称焦炭[i] / 工艺称焦炭[5];
              //    落地湿焦[i] = 落地湿焦[5] * 工艺称焦炭[i] / 工艺称焦炭[5];
              //}
            }
            机烧[5] = 机烧[0] + 机烧[1] + 机烧[2] + 机烧[3] + 机烧[4];
            //干毛焦[5] = 自产湿焦[5] * (1 - 自产湿焦水分 / 100) + 落地湿焦[5] * (1 - 落地湿焦水分 / 100);
            for (int i = 0; i < 5; i++)
            {
                干毛焦[i] = 自产湿焦[i] * (1 - 自产湿焦水分 / 100) + 落地湿焦[i] * (1 - 落地湿焦水分 / 100);
                入炉焦炭[i] = 干毛焦[i] - 湿焦粉[i] * (1 - 湿焦粉水分 / 100);
            }
            干毛焦[5] = 0;
            入炉焦炭[5] = 0;
            for (int i = 0; i < 5; i++)
            {
                干毛焦[5] += 干毛焦[i];
                入炉焦炭[5] += 入炉焦炭[i];
            }
            //入炉焦炭[5] = 干毛焦[5] - 湿焦粉[5] * (1 - 湿焦粉水分 / 100);
            //for (int i = 0; i < 5; i++)
            //{
            //    if (工艺称焦炭[5] > 0)
            //    {  
            //        入炉焦炭[i] = 入炉焦炭[5] * 工艺称焦炭[i] / 工艺称焦炭[5];      
            //    }
            //    干毛焦[i] = 入炉焦炭[i] + 湿焦粉[i] * (1 - 湿焦粉水分 / 100);
            //}
            for (int i=0;i<6;i++)
            {
                总耗[i] = 机烧[i] + 竖球[i] + 熟料[i] + 生料[i] + 本溪矿[i];
                if (总耗[i] > 0)
                    熟料比[i] = (机烧[i] + 竖球[i] + 熟料[i]) * 100 / 总耗[i];
                综合焦炭[i] = 入炉焦炭[i] + 0.8 * 煤粉[i] + 0.9 * 焦丁[i];

                if (冷风流量[i] > 0)
                    富氧率[i] = 0.79 * 富氧量[i]*100 / 冷风流量[i] / 24;
                if (入炉焦炭[i] > 0)
                    焦炭负荷[i] = 总耗[i] / 入炉焦炭[i];
            }

            冶炼强度[0] = 入炉焦炭[0] / 有效炉容[0];
            冶炼强度[1] = 入炉焦炭[1] / 有效炉容[1];
            冶炼强度[2] = 入炉焦炭[2] / 有效炉容[2];
            冶炼强度[3] = 入炉焦炭[3] / 有效炉容[3];
            冶炼强度[4] = 入炉焦炭[4] / 有效炉容[4];
            冶炼强度[5] = 入炉焦炭[5] / 有效炉容[5];
            
            for (int i = 0; i < 6; i++)
            {
                总耗[i] = double.Parse(总耗[i].ToString("#######0"));
                机烧[i] = double.Parse(机烧[i].ToString("#######0"));
                竖球[i] = double.Parse(竖球[i].ToString("#######0"));
                生料[i] = double.Parse(生料[i].ToString("#######0"));
                熟料[i] = double.Parse(熟料[i].ToString("#######0"));
                本溪矿[i] = double.Parse(本溪矿[i].ToString("#######0"));
                自产湿焦[i] = double.Parse(自产湿焦[i].ToString("#######0"));
                落地湿焦[i] = double.Parse(落地湿焦[i].ToString("#######0"));
                干毛焦[i] = double.Parse(干毛焦[i].ToString("#######0.00"));
                湿焦粉[i] = double.Parse(湿焦粉[i].ToString("#######0"));
                入炉焦炭[i] = double.Parse(入炉焦炭[i].ToString("#######0.0"));
                煤粉[i] = double.Parse(煤粉[i].ToString("#######0.0"));
                焦丁[i] = double.Parse(焦丁[i].ToString("#######0.0"));
                综合焦炭[i] = double.Parse(综合焦炭[i].ToString("#######0.0"));
                熟料比[i] = double.Parse(熟料比[i].ToString("#######0.00"));
                富氧率[i] = double.Parse(富氧率[i].ToString("#######0.00"));
            }

            //累计
            //大中修
            cmd.CommandText = "select 高炉,sum(间隔) from 休风 where trunc(时间,'MONTH')=trunc(:rq,'MONTH') and trunc(时间)<=trunc(:rq) and (分类='大修' or 分类='中修')  and  间隔>0　and 间隔 is not null group by 高炉";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                int gaolu = dr.IsDBNull(0) ? 6 : dr.GetInt32(0);
                if (gaolu < 6) 累计大中修[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
            }
            累计大中修[5] = 累计大中修[0] + 累计大中修[1] + 累计大中修[2] + 累计大中修[3] + 累计大中修[4];
            dr.Close();

            cmd.CommandText = "select SUM(P05二号皮带),SUM(P06三号皮带),SUM(P07总返矿),SUM(自产湿焦),SUM(落地湿焦),TRUNC(AVG(P02自产焦水分),2),TRUNC(AVG(P03落地焦水分),2),TRUNC(AVG(P04焦粉水分),2),TRUNC(SUM(NVL(自产湿焦,0)*(1-NVL(P02自产焦水分,0)/100)+NVL(落地湿焦,0)*(1-NVL(P03落地焦水分,0)/100)),2) as 累计干毛焦 from 全厂日消耗 where trunc(P01日期,'MONTH')=trunc(:rq,'MONTH') and P01日期<=:rq";
            dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                累计二号皮带 = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
                累计三号皮带 = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                累计返矿 = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                //累计自产湿焦[5] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                //累计落地湿焦[5] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                累计自产湿焦水分 = dr.IsDBNull(5) ? 0 : dr.GetDouble(5);
                累计落地湿焦水分 = dr.IsDBNull(6) ? 0 : dr.GetDouble(6);
                累计湿焦粉水分 = dr.IsDBNull(7) ? 0 : dr.GetDouble(7);
                累计干毛焦[5]=dr.IsDBNull(8) ? 0 : dr.GetDouble(8);

            }
            dr.Close();

            cmd.CommandText = "select 高炉,SUM(trunc(机烧矿,2)),SUM(trunc(竖球,2)),SUM(trunc(熟料,2)),SUM(trunc(生料,2)),SUM(trunc(本溪矿,2)),SUM(工艺称),SUM(煤粉),SUM(焦丁),TRUNC(SUM(自产湿焦),2),TRUNC(SUM(落地湿焦),2) from 原料消耗 where  trunc(日期,'MONTH')=trunc(:rq,'MONTH') and 日期<=:rq group by 高炉";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    累计机烧[gaolu - 1] = dr.IsDBNull(1) ? 0 : Convert.ToInt32(dr.GetDouble(1));
                    累计竖球[gaolu - 1] = dr.IsDBNull(2) ? 0 : Convert.ToInt32(dr.GetDouble(2));
                    累计熟料[gaolu - 1] = dr.IsDBNull(3) ? 0 : Convert.ToInt32(dr.GetDouble(3));
                    累计生料[gaolu - 1] = dr.IsDBNull(4) ? 0 : Convert.ToInt32(dr.GetDouble(4));
                    累计本溪矿[gaolu - 1] = dr.IsDBNull(5) ? 0 : Convert.ToInt32(dr.GetDouble(5));
                    累计工艺称焦炭[gaolu - 1] = dr.IsDBNull(6) ? 0 : dr.GetDouble(6);
                    累计煤粉[gaolu - 1] = dr.IsDBNull(7) ? 0 : dr.GetDouble(7);
                    累计焦丁[gaolu - 1] = dr.IsDBNull(8) ? 0 : dr.GetDouble(8);
                    累计自产湿焦[gaolu-1] = dr.IsDBNull(9) ? 0 : dr.GetDouble(9);
                    累计落地湿焦[gaolu-1] = dr.IsDBNull(10) ? 0 : dr.GetDouble(10);
                }
            }
            dr.Close();

            cmd.CommandText = "select gaolu,sum(shijiaofen),sum(fuyang),sum(lengfengliuliang),sum(feitie),trunc(avg(fengwen),0),trunc(avg(ludingwendu),0),trunc(avg(refengyali),0),trunc(avg(ludingyali),0),sum(fengkoudatao),sum(fengkouzhongtao),sum(fengkouxiaotao),sum(zhakoudatao),sum(zhakouzhongtao),sum(zhakouxiaotao),sum(zuoliao),sum(xuanliao),sum(bengliao),trunc(sum(shijiaofen*(1-全厂日消耗.P04焦粉水分/100)),2)  from xiaohao,全厂日消耗 where 全厂日消耗.P01日期=xiaohao.RQ and trunc(RQ,'MONTH')=trunc(:rq,'MONTH') and RQ<=:rq  group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    累计湿焦粉[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    累计富氧量[gaolu - 1] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                    累计冷风流量[gaolu - 1] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                    累计废铁[gaolu - 1] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                    累计平均风温[gaolu - 1] = dr.IsDBNull(5) ? 0 : dr.GetDouble(5);
                    累计炉顶温度[gaolu - 1] = dr.IsDBNull(6) ? 0 : dr.GetDouble(6);
                    累计热风压力[gaolu - 1] = dr.IsDBNull(7) ? 0 : dr.GetDouble(7);
                    累计炉顶压力[gaolu - 1] = dr.IsDBNull(8) ? 0 : dr.GetDouble(8);
                    累计风口损坏数大[gaolu - 1] = dr.IsDBNull(9) ? 0 : dr.GetDouble(9);
                    累计风口损坏数中[gaolu - 1] = dr.IsDBNull(10) ? 0 : dr.GetDouble(10);
                    累计风口损坏数小[gaolu - 1] = dr.IsDBNull(11) ? 0 : dr.GetDouble(11);
                    累计渣口损坏数大[gaolu - 1] = dr.IsDBNull(12) ? 0 : dr.GetDouble(12);
                    累计渣口损坏数中[gaolu - 1] = dr.IsDBNull(13) ? 0 : dr.GetDouble(13);
                    累计渣口损坏数小[gaolu - 1] = dr.IsDBNull(14) ? 0 : dr.GetDouble(14);
                    累计坐料次数[gaolu - 1] = dr.IsDBNull(15) ? 0 : dr.GetDouble(15);
                    累计悬料次数[gaolu - 1] = dr.IsDBNull(16) ? 0 : dr.GetDouble(16);
                    累计崩料次数[gaolu - 1] = dr.IsDBNull(17) ? 0 : dr.GetDouble(17);
                    累计干焦粉[gaolu - 1] = dr.IsDBNull(18) ? 0 : dr.GetDouble(18);
                }
            }
            dr.Close();
            cmd.CommandText = "select 高炉,sum(间隔) from 休风 where   trunc(时间,'MONTH')=trunc(:rq,'MONTH') and trunc(时间)<=:rq group by 高炉";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    累计休风情况[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
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
                    累计慢风[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                }
            }
            dr.Close();
            
            for (int i = 0; i < 5; i++)
            {
                累计休风情况[5] += 累计休风情况[i];
                累计慢风[5] += 累计慢风[i];
                累计坐料次数[5] += 累计坐料次数[i];
                累计悬料次数[5] += 累计悬料次数[i];
                累计崩料次数[5] += 累计崩料次数[i];
                累计风口损坏数大[5] += 累计风口损坏数大[i];
                累计风口损坏数中[5] += 累计风口损坏数中[i];
                累计风口损坏数小[5] += 累计风口损坏数小[i];
                累计渣口损坏数大[5] += 累计渣口损坏数大[i];
                累计渣口损坏数中[5] += 累计渣口损坏数中[i];
                累计渣口损坏数小[5] += 累计渣口损坏数小[i];
                累计废铁[5] += 累计废铁[i];
                累计机烧[5] += 累计机烧[i];
                累计竖球[5] += 累计竖球[i];
                累计熟料[5] += 累计熟料[i];
                累计生料[5] += 累计生料[i];
                累计本溪矿[5] += 累计本溪矿[i];
                累计工艺称焦炭[5] += 累计工艺称焦炭[i];
                累计湿焦粉[5] += 累计湿焦粉[i];
                累计煤粉[5] += 累计煤粉[i];
                累计焦丁[5] += 累计焦丁[i];
                累计富氧量[5] += 累计富氧量[i];
                累计冷风流量[5] += 累计冷风流量[i];
                累计热风压力[5] += 累计热风压力[i];
                累计炉顶温度[5] += 累计炉顶温度[i];
                累计炉顶压力[5] += 累计炉顶压力[i];
                //累计干焦粉[5] += 累计干焦粉[i];
                累计自产湿焦[5] += 累计自产湿焦[i];
                累计落地湿焦[5] += 累计落地湿焦[i];

            }
            累计有效炉容[0] = (((1440 * rq.Day) - 累计大中修[0]) / (1440 * rq.Day)) * Constants.volume1 * rq.Day;
            累计有效炉容[1] = (((1440 * rq.Day) - 累计大中修[1]) / (1440 * rq.Day)) * Constants.volume2 * rq.Day;
            累计有效炉容[2] = (((1440 * rq.Day) - 累计大中修[2]) / (1440 * rq.Day)) * Constants.volume3 * rq.Day;
            累计有效炉容[3] = (((1440 * rq.Day) - 累计大中修[3]) / (1440 * rq.Day)) * Constants.volume4 * rq.Day;
            累计有效炉容[4] = (((1440 * rq.Day) - 累计大中修[4]) / (1440 * rq.Day)) * Constants.volume5 * rq.Day;
            累计有效炉容[5] = 累计有效炉容[0] + 累计有效炉容[1] + 累计有效炉容[2] + 累计有效炉容[3] + 累计有效炉容[4];

            累计平均风温[5] = (累计平均风温[0] * Convert.ToInt32(累计平均风温[0] > 0) + 累计平均风温[1] * Convert.ToInt32(累计平均风温[1] > 0) + 累计平均风温[2] * Convert.ToInt32(累计平均风温[2] > 0) + 累计平均风温[3] * Convert.ToInt32(累计平均风温[3] > 0) + 累计平均风温[4] * Convert.ToInt32(累计平均风温[4] > 0)) / (Convert.ToInt32(累计平均风温[0] > 0) + Convert.ToInt32(累计平均风温[1] > 0) + Convert.ToInt32(累计平均风温[2] > 0) + Convert.ToInt32(累计平均风温[3] > 0) + Convert.ToInt32(累计平均风温[4] > 0));
            累计热风压力[5] = (累计热风压力[0] * Convert.ToInt32(累计热风压力[0] > 0) + 累计热风压力[1] * Convert.ToInt32(累计热风压力[1] > 0) + 累计热风压力[2] * Convert.ToInt32(累计热风压力[2] > 0) + 累计热风压力[3] * Convert.ToInt32(累计热风压力[3] > 0) + 累计热风压力[4] * Convert.ToInt32(累计热风压力[4] > 0)) / (Convert.ToInt32(累计热风压力[0] > 0) + Convert.ToInt32(累计热风压力[1] > 0) + Convert.ToInt32(累计热风压力[2] > 0) + Convert.ToInt32(累计热风压力[3] > 0) + Convert.ToInt32(累计热风压力[4] > 0));
            累计炉顶温度[5] = (累计炉顶温度[0] * Convert.ToInt32(累计炉顶温度[0] > 0) + 累计炉顶温度[1] * Convert.ToInt32(累计炉顶温度[1] > 0) + 累计炉顶温度[2] * Convert.ToInt32(累计炉顶温度[2] > 0) + 累计炉顶温度[3] * Convert.ToInt32(累计炉顶温度[3] > 0) + 累计炉顶温度[4] * Convert.ToInt32(累计炉顶温度[4] > 0)) / (Convert.ToInt32(累计炉顶温度[0] > 0) + Convert.ToInt32(累计炉顶温度[1] > 0) + Convert.ToInt32(累计炉顶温度[2] > 0) + Convert.ToInt32(累计炉顶温度[3] > 0) + Convert.ToInt32(累计炉顶温度[4] > 0));
            累计炉顶压力[5] = (累计炉顶压力[0] * Convert.ToInt32(累计炉顶压力[0] > 0) + 累计炉顶压力[1] * Convert.ToInt32(累计炉顶压力[1] > 0) + 累计炉顶压力[2] * Convert.ToInt32(累计炉顶压力[2] > 0) + 累计炉顶压力[3] * Convert.ToInt32(累计炉顶压力[3] > 0) + 累计炉顶压力[4] * Convert.ToInt32(累计炉顶压力[4] > 0)) / (Convert.ToInt32(累计炉顶压力[0] > 0) + Convert.ToInt32(累计炉顶压力[1] > 0) + Convert.ToInt32(累计炉顶压力[2] > 0) + Convert.ToInt32(累计炉顶压力[3] > 0) + Convert.ToInt32(累计炉顶压力[4] > 0));

           // 累计平均风温[5] = (累计平均风温[0] * 350 + 累计平均风温[1] * 350 + 累计平均风温[2] * 350 + 累计平均风温[3] * 400 + 累计平均风温[4] * 420) / 1870;
           // 累计热风压力[5] = 累计热风压力[5] / 5;
           // 累计炉顶温度[5] = 累计炉顶温度[5] / 5;
          //  累计炉顶压力[5] = 累计炉顶压力[5] / 5;

            累计全厂机烧 = 累计二号皮带 + 累计三号皮带 - 累计返矿 + ((累计竖球[5] + 累计熟料[5] + 累计生料[5] + 累计本溪矿[5]) / 0.92 - (累计竖球[5] + 累计熟料[5] + 累计生料[5] + 累计本溪矿[5]));

            for (int i = 0; i < 5; i++)
            {
                if (累计机烧[5] > 0)
                    累计机烧[i] = double.Parse(((double)(累计全厂机烧 * 累计机烧[i] / 累计机烧[5])).ToString("########0")); //分配机烧

                if (累计工艺称焦炭[5] > 0)
                {
                    累计自产湿焦[i] = 累计自产湿焦[5] * 累计工艺称焦炭[i] / 累计工艺称焦炭[5];
                    累计落地湿焦[i] = 累计落地湿焦[5] * 累计工艺称焦炭[i] / 累计工艺称焦炭[5];
                }
            }
            累计机烧[5] = 累计机烧[0] + 累计机烧[1] + 累计机烧[2] + 累计机烧[3] + 累计机烧[4];

            //for (int i = 0; i < 5; i++)
            //{
            //    累计干毛焦[i] = 累计自产湿焦[i] * (1 - 累计自产湿焦水分 / 100) + 累计落地湿焦[i] * (1 - 累计落地湿焦水分 / 100);
            //    累计入炉焦炭[i] = 累计干毛焦[i] - 累计湿焦粉[i] * (1 - 累计湿焦粉水分 / 100);
            //}
            //累计干毛焦[5] = 0;
            //累计入炉焦炭[5] = 0;
            //for (int i = 0; i < 5; i++)
            //{
            //    累计干毛焦[5] += 累计干毛焦[i];
            //    累计入炉焦炭[5] += 累计入炉焦炭[i];
            //}
            if (rq.Day == 1)
            {
                for (int i = 0; i < 6; i++)
                {
                    累计干毛焦[i] = 干毛焦[i];
                    累计入炉焦炭[i] = 入炉焦炭[i];
                }
            }
            else
            {
                DateTime rq2 = rq.AddDays(-1);

                cmd.CommandText = "select P01单位,P38干毛焦,P27入炉焦炭总耗 from 技术日报 where trunc(P日期)=trunc(:rq) and P02项目='累计'";
                cmd.Parameters.Add(":rq", OracleType.DateTime).Value = rq2;
                dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    int gaolu = 7;
                    string gaolustr = dr.GetString(0);
                    switch (gaolustr)
                    {
                        case "1#":
                            gaolu = 0;
                            break;
                        case "2#":
                            gaolu = 1;
                            break;
                        case "3#":
                            gaolu = 2;
                            break;
                        case "4#":
                            gaolu = 3;
                            break;
                        case "5#":
                            gaolu = 4;
                            break;
                        case "Q":
                            gaolu = 5;
                            break;
                    }
                    if (gaolu < 6)
                    {
                        累计干毛焦[gaolu] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                        累计入炉焦炭[gaolu] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                        累计干毛焦[gaolu] += 干毛焦[gaolu];
                        累计入炉焦炭[gaolu] += 入炉焦炭[gaolu];
                    }
                }
                dr.Close();
            }
            cn.Close();
            //累计干毛焦[5] = 累计自产湿焦[5] * (1 - 累计自产湿焦水分 / 100) + 累计落地湿焦[5] * (1 - 累计落地湿焦水分 / 100);
            //累计入炉焦炭[5] = 累计干毛焦[5] - 累计干焦粉[5];
            //for (int i = 0; i < 5; i++)
            //{
            //    if (累计工艺称焦炭[5] > 0)
            //    {
            //        累计入炉焦炭[i] = 累计入炉焦炭[5] * 累计工艺称焦炭[i] / 累计工艺称焦炭[5];
            //    }
            //    累计干毛焦[i] = 累计入炉焦炭[i] + 累计湿焦粉[i] * (1 - 累计湿焦粉水分 / 100);
            //}

            for (int i = 0; i < 6; i++)
            {
                累计总耗[i] = 累计机烧[i] + 累计竖球[i] + 累计熟料[i] + 累计生料[i] + 累计本溪矿[i];
                if (累计总耗[i] > 0)
                    累计熟料比[i] = (累计机烧[i] + 累计竖球[i] + 累计熟料[i]) * 100 / 累计总耗[i];
                累计综合焦炭[i] = 累计入炉焦炭[i] + 0.8 * 累计煤粉[i] + 0.9 * 累计焦丁[i];

                if (累计冷风流量[i] > 0)
                    累计富氧率[i] = 0.79 * 累计富氧量[i]*100 / 累计冷风流量[i] / 24;
                if (累计入炉焦炭[i] > 0)
                    累计焦炭负荷[i] = 累计总耗[i] / 累计入炉焦炭[i];
            }
            累计冶炼强度[0] = 累计入炉焦炭[0] / 累计有效炉容[0];
            累计冶炼强度[1] = 累计入炉焦炭[1] / 累计有效炉容[1];
            累计冶炼强度[2] = 累计入炉焦炭[2] / 累计有效炉容[2];
            累计冶炼强度[3] = 累计入炉焦炭[3] / 累计有效炉容[3];
            累计冶炼强度[4] = 累计入炉焦炭[4] / 累计有效炉容[4];
            累计冶炼强度[5] = 累计入炉焦炭[5] / 累计有效炉容[5];

            for (int i = 0; i < 6; i++)
            {
                累计总耗[i] = double.Parse(累计总耗[i].ToString("#######0"));
                累计机烧[i] = double.Parse(累计机烧[i].ToString("#######0"));
                累计竖球[i] = double.Parse(累计竖球[i].ToString("#######0"));
                累计生料[i] = double.Parse(累计生料[i].ToString("#######0"));
                累计熟料[i] = double.Parse(累计熟料[i].ToString("#######0"));
                累计本溪矿[i] = double.Parse(累计本溪矿[i].ToString("#######0"));
                累计自产湿焦[i] = double.Parse(累计自产湿焦[i].ToString("#######0"));
                累计落地湿焦[i] = double.Parse(累计落地湿焦[i].ToString("#######0"));
                累计干毛焦[i] = double.Parse(累计干毛焦[i].ToString("#######0.00"));
                累计湿焦粉[i] = double.Parse(累计湿焦粉[i].ToString("#######0"));
                累计入炉焦炭[i] = double.Parse(累计入炉焦炭[i].ToString("#######0.0"));
                累计煤粉[i] = double.Parse(累计煤粉[i].ToString("#######0.0"));
                累计焦丁[i] = double.Parse(累计焦丁[i].ToString("#######0.0"));
                累计综合焦炭[i] = double.Parse(累计综合焦炭[i].ToString("#######0.0"));
                累计熟料比[i] = double.Parse(累计熟料比[i].ToString("#######0.00"));
                累计富氧率[i] = double.Parse(累计富氧率[i].ToString("#######0.00"));
            }
            cn.Close();
        }
    }
}
