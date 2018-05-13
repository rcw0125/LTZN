using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;

namespace LTZN.技术月报
{

    class 上月累计
    {
        public int 月数 = 0;
        public int 累计天数 = 30;
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
        public double[] 干焦粉 = new double[6];
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

        public void getDataBy(int nian, int yue)
        {

            月数 = yue;
            DateTime d = new DateTime(nian, yue + 1, 1);
            累计天数 = d.DayOfYear - 1;
            累计天数 += DateTime.DaysInMonth(nian, yue + 1);
            OracleConnection cn = new OracleConnection();
            cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
            cn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = cn;

            cmd.Parameters.Add(":nian", OracleType.Int32).Value = nian;
            cmd.Parameters.Add(":yue", OracleType.Int32).Value = yue;

            cmd.CommandText = "select P01单位,P03全铁产量,P04合格铁,P05制钢铁,P07折合产量,一级铁,优质铁,P65煤粉总耗,P66焦丁总耗"
                               + ",P61干毛焦,P63入炉焦炭总耗,P60七号称,P50机烧消耗,P51竖炉球消耗,P53其它熟料消耗,P54本溪矿消耗,P55其它生料消耗,P56废铁总耗,瓦斯灰"
                               + ",P40风口大套损坏数,P41风口中套损坏数,P42风口小套损坏数,P43渣口大套损坏数,P44渣口中套损坏数,P45渣口小套损坏数"
                               + ",P37悬料次数,P38坐料次数,P39崩料次数,P20休风时间,P22慢风时间,P29冷风流量,富氧量"
                               + ",收入含铁,支出含铁,高压操作时间,有效工作时间,理论渣量,深料线,全部料线,高炉有效容积,高炉实际容积 "
                               + ",P69铁成分SI,P70铁成分MN,P71铁成分P,P72铁成分S,P73渣成分CAO,P74渣成分SIO2,P75渣成分AL2O3,P76渣成分MGO,P78渣成分S,P80含SI偏差制铁量,P81含SI偏差铸造铁,P32热风压力,P34炉顶温度,P33炉顶压力"
                               + ",P30平均风温,P24煤气成分CO2,P25煤气成分计算1,入炉矿含铁 "
                               + " from 技术月报 where P年=:nian and P月=:yue and P02项目='累计'";
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
                    全铁产量[i] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    合格铁[i] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                    制钢铁[i] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                    折算产量[i] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                    一级铁[i] = dr.IsDBNull(5) ? 0 : dr.GetDouble(5);
                    送炼钢优质铁[i] = dr.IsDBNull(6) ? 0 : dr.GetDouble(6);
                    煤粉[i] = dr.IsDBNull(7) ? 0 : dr.GetDouble(7);
                    焦丁[i] = dr.IsDBNull(8) ? 0 : dr.GetDouble(8);
                    干毛焦[i] = dr.IsDBNull(9) ? 0 : dr.GetDouble(9);
                    入炉焦炭[i] = dr.IsDBNull(10) ? 0 : dr.GetDouble(10);
                    七号称焦炭[i] = dr.IsDBNull(11) ? 0 : dr.GetDouble(11);
                    机烧[i] = dr.IsDBNull(12) ? 0 : dr.GetDouble(12);
                    竖球[i] = dr.IsDBNull(13) ? 0 : dr.GetDouble(13);
                    其它熟料[i] = dr.IsDBNull(14) ? 0 : dr.GetDouble(14);
                    本溪矿[i] = dr.IsDBNull(15) ? 0 : dr.GetDouble(15);
                    其它生料[i] = dr.IsDBNull(16) ? 0 : dr.GetDouble(16);
                    废铁[i] = dr.IsDBNull(17) ? 0 : dr.GetDouble(17);
                    瓦斯灰[i] = dr.IsDBNull(18) ? 0 : dr.GetDouble(18);
                    风口损坏大[i] = dr.IsDBNull(19) ? 0 : dr.GetDouble(19);
                    风口损坏中[i] = dr.IsDBNull(20) ? 0 : dr.GetDouble(20);
                    风口损坏小[i] = dr.IsDBNull(21) ? 0 : dr.GetDouble(21);
                    渣口损坏大[i] = dr.IsDBNull(22) ? 0 : dr.GetDouble(22);
                    渣口损坏中[i] = dr.IsDBNull(23) ? 0 : dr.GetDouble(23);
                    渣口损坏小[i] = dr.IsDBNull(24) ? 0 : dr.GetDouble(24);
                    悬料次数[i] = dr.IsDBNull(25) ? 0 : dr.GetDouble(25);
                    坐料次数[i] = dr.IsDBNull(26) ? 0 : dr.GetDouble(26);
                    崩料次数[i] = dr.IsDBNull(27) ? 0 : dr.GetDouble(27);
                    休风小时[i] = dr.IsDBNull(28) ? 0 : dr.GetDouble(28);
                    慢风[i] = dr.IsDBNull(29) ? 0 : dr.GetDouble(29);
                    冷风流量[i] = dr.IsDBNull(30) ? 0 : dr.GetDouble(30);
                    富氧流量[i] = dr.IsDBNull(31) ? 0 : dr.GetDouble(31);
                    收入含铁[i] = dr.IsDBNull(32) ? 0 : dr.GetDouble(32);
                    支出含铁[i] = dr.IsDBNull(33) ? 0 : dr.GetDouble(33);
                    高压操作时间[i] = dr.IsDBNull(34) ? 0 : dr.GetDouble(34);
                    有效工作时间[i] = dr.IsDBNull(35) ? 0 : dr.GetDouble(35);
                    理论渣量[i] = dr.IsDBNull(36) ? 0 : dr.GetDouble(36);
                    深料线[i] = dr.IsDBNull(37) ? 0 : dr.GetDouble(37);
                    全部料线[i] = dr.IsDBNull(38) ? 0 : dr.GetDouble(38);
                    高炉有效炉容[i] = dr.IsDBNull(39) ? 0 : dr.GetDouble(39);
                    高炉实际炉容[i] = dr.IsDBNull(40) ? 0 : dr.GetDouble(40);
                    铁Si[i] = dr.IsDBNull(41) ? 0 : dr.GetDouble(41);
                    铁Mn[i] = dr.IsDBNull(42) ? 0 : dr.GetDouble(42);
                    铁P[i] = dr.IsDBNull(43) ? 0 : dr.GetDouble(43);
                    铁S[i] = dr.IsDBNull(44) ? 0 : dr.GetDouble(44);
                    渣CaO[i] = dr.IsDBNull(45) ? 0 : dr.GetDouble(45);
                    渣SiO2[i] = dr.IsDBNull(46) ? 0 : dr.GetDouble(46);
                    渣Al2O3[i] = dr.IsDBNull(47) ? 0 : dr.GetDouble(47);
                    渣MgO[i] = dr.IsDBNull(48) ? 0 : dr.GetDouble(48);
                    渣S[i] = dr.IsDBNull(49) ? 0 : dr.GetDouble(49);
                    制钢铁Si偏差[i] = dr.IsDBNull(50) ? 0 : dr.GetDouble(50);
                    铸造铁Si偏差[i] = dr.IsDBNull(51) ? 0 : dr.GetDouble(51);
                    热风压力[i] = dr.IsDBNull(52) ? 0 : dr.GetDouble(52);
                    炉顶温度[i] = dr.IsDBNull(53) ? 0 : dr.GetDouble(53);
                    炉顶压力[i] = dr.IsDBNull(54) ? 0 : dr.GetDouble(54);
                    平均风温[i] = dr.IsDBNull(55) ? 0 : dr.GetDouble(55);
                    CO2[i] = dr.IsDBNull(56) ? 0 : dr.GetDouble(56);
                    CO2率[i] = dr.IsDBNull(57) ? 0 : dr.GetDouble(57);
                    入炉矿含铁[i] = dr.IsDBNull(58) ? 0 : dr.GetDouble(58);
                }
            }
            dr.Close();


            cmd.CommandText = "select 高炉,sum(工艺称) from 原料消耗 where extract(year from 日期)=:nian and extract(month from 日期)<=:yue group by 高炉";
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

            cn.Close();
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
            高炉有效炉容[0] = Constants.volume1 * (累计天数 - 休风[0] / 1440);
            高炉有效炉容[1] = Constants.volume2 * (累计天数 - 休风[1] / 1440);
            高炉有效炉容[2] = Constants.volume3 * (累计天数 - 休风[2] / 1440);
            高炉有效炉容[3] = Constants.volume4 * (累计天数 - 休风[3] / 1440);
            高炉有效炉容[4] = Constants.volume5 * (累计天数 - 休风[4] / 1440);
            高炉有效炉容[5] = 高炉有效炉容[0] + 高炉有效炉容[1] + 高炉有效炉容[2] + 高炉有效炉容[3] + 高炉有效炉容[4];


        }
        private void 计算铸造铁()
        {
            for (int i = 0; i < 6; i++)
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
                合格率[i] = double.Parse(合格率[i].ToString("N2"));
            }
        }
        private void 计算一级品率()
        {
            for (int i = 0; i < 6; i++)
            {
                if (全铁产量[i] > 0)
                    一级品率[i] = 一级铁[i] * 100 / 全铁产量[i];
                一级品率[i] = double.Parse(一级品率[i].ToString("N2"));
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
            高炉实际炉容[0] = Constants.volume1 * (累计天数-大中休[0]/24);
            高炉实际炉容[1] = Constants.volume2 * (累计天数-大中休[1]/24);
            高炉实际炉容[2] = Constants.volume3 * (累计天数 - 大中休[2] / 24);
            高炉实际炉容[3] = Constants.volume4 * (累计天数 - 大中休[3] / 24);
            高炉实际炉容[4] = Constants.volume5 * (累计天数 - 大中休[4] / 24);
            高炉实际炉容[5] = 高炉有效炉容[0] + 高炉有效炉容[1] + 高炉有效炉容[2] + 高炉有效炉容[3] + 高炉有效炉容[4];
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
        private void 计算入炉焦炭冶炼强度()
        {
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
        private void 计算折算综合焦比()
        {
            for (int i = 0; i < 6; i++)
            {
                if (折算产量[i] > 0)
                    折算综合焦比[i] = 综合焦炭[i] * 1000 / 折算产量[i];
                折算综合焦比[i] = double.Parse(折算综合焦比[i].ToString("N0"));
            }
        }
        private void 计算折合入炉焦比()
        {
            for (int i = 0; i < 6; i++)
            {
                if (折算产量[i] > 0)
                    折合入炉焦比[i] = 入炉焦炭[i] * 1000 / 折算产量[i];
                折合入炉焦比[i] = double.Parse(折合入炉焦比[i].ToString("N0"));
            }
        }
        private void 计算入炉矿含铁()
        {
            for (int i = 0; i < 6; i++)
            {
                入炉矿含铁[i] = 机烧[i] * 机烧品位 / 100 + 竖球[i] * 竖球品位 / 100 + 其它熟料[i] * 其它熟料品位 / 100 + 本溪矿[i] * 本溪矿品位 / 100 + 其它生料[i] * 其它生料品位 / 100;
            }
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
        private void 计算熟料比()
        {
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
            
            for (int i = 0; i < 6; i++)
            {
                if (有效工作时间[i] > 0)
                {
                    休风率[i] = 休风小时[i] * 100 / 有效工作时间[i];
                    休风率[i] = double.Parse(休风率[i].ToString("N2"));
                }
                else
                    休风率[i] = 0;
            }
        }
        private void 计算慢风率()
        {
            for (int i = 0; i < 6; i++)
            {
                if (有效工作时间[i] > 0)
                {
                    慢风率[i] = 慢风[i] * 100 / 有效工作时间[i];
                    慢风率[i] = double.Parse(慢风率[i].ToString("N2"));
                }
                else
                    慢风率[i] = 0;
            }
        }
        private void 计算CO2率()
        {
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
        private void 计算回收率()
        {
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
                   // 富氧率[i] = 0.79 * 富氧流量[i] * 100 / 冷风流量[i];
                富氧率[i] = double.Parse(富氧率[i].ToString("N2"));
            }
        }
        private void 计算有效工作时间()
        {
            for (int i = 0; i < 5; i++)
            {
                有效工作时间[i] = 累计天数 * 24;
            }
            有效工作时间[5] = 累计天数 * 5 * 24;

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
                    理论渣量[i] = (机烧[i] * 机烧CaO + 竖球[i] * 竖球CaO + 其它熟料[i] * 其它熟料CaO + 本溪矿[i] * 本溪矿CaO + 其它生料[i] * 其它生料CaO + 入炉焦炭[i] * 焦炭灰份 * 1 + 煤粉[i] * 煤粉灰份 * 2) / 渣CaO[5];
            }
        }
        private void 计算渣铁比()
        {
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
                    灰铁比[i] = 瓦斯灰[i] * 1000 / 全铁产量[i];
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
        private void 计算废铁单耗()
        {
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
                收入含铁[i] = 机烧[i] * 机烧品位 + 竖球[i] * 竖球品位 + 其它熟料[i] * 其它熟料品位 + 本溪矿[i] * 本溪矿品位 + 其它生料[i] * 其它生料品位 + 0.9 * 废铁[i];
            }
        }
        private void 计算支出含铁()
        {
            for (int i = 0; i < 6; i++)
            {
                支出含铁[i] = 全铁产量[i] * 铁Fe[i];
            }
        }
        private void 计算金属收得率()
        {
            for (int i = 0; i < 6; i++)
            {
                if (收入含铁[i] > 0)
                    金属收得率[i] = 支出含铁[i] * 100 / 收入含铁[i];
                金属收得率[i] = double.Parse(金属收得率[i].ToString("N2"));
            }

        }
        private void 计算综合焦炭单耗()
        {
            for (int i = 0; i < 6; i++)
            {
                if (合格铁[i] > 0)
                    综合焦炭单耗[i] = 综合焦炭[i] * 1000 / 合格铁[i];
                综合焦炭单耗[i] = double.Parse(综合焦炭单耗[i].ToString("N0"));
            }
        }
        private void 计算入炉焦炭单耗()
        {
            for (int i = 0; i < 6; i++)
            {
                if (合格铁[i] > 0)
                    入炉焦炭单耗[i] = 入炉焦炭[i] * 1000 / 合格铁[i];
                入炉焦炭单耗[i] = double.Parse(入炉焦炭单耗[i].ToString("N0"));
            }
        }
        private void 计算焦丁单耗()
        {
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
        private void 计算燃料比()
        {
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
                    工艺称焦比[i] = 工艺称焦炭[i] * 干毛焦[i] *1000/ 七号称焦炭[i] / 全铁产量[i];
                工艺称焦比[i] = double.Parse(工艺称焦比[i].ToString("N0"));
            }
        }
        public void 计算()
        {
            计算铸造铁();
            计算干焦粉();
            计算原料总耗();
            计算综合焦炭();
           // 计算高炉有效炉容();
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
            //计算CO2率();
            计算深料线率();
            计算回收率();
            计算富氧率();
            计算高压率();
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

