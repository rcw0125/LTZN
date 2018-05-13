using System;
using System.Collections.Generic;
using System.Data.OracleClient;
using System.Text;
using LTZN.技术日报;

namespace LTZN.Repository
{
    public class RiReportRepository : RepositoryBase
    {
        public 技术日报内容 GetRiReport(DateTime riqi)
        {
            技术日报内容 result = new 技术日报内容(riqi);

            StringBuilder sbSql = new StringBuilder();
            sbSql.AppendLine("SELECT P01单位, P02项目, P03合格铁, P04炼钢铁, P05铸造铁, P06号外铁,");
            sbSql.AppendLine("P07合格率, P08高炉利用系数, P09高炉实物系数, P10原料矿合计总耗, ");
            sbSql.AppendLine("P11原料矿合计单耗, P12原料矿机烧, P13原料矿竖炉球, P14原料矿CT, ");
            sbSql.AppendLine("P15原料矿其它熟料, P16原料矿本溪矿, P17原料矿其它生料, P18废铁总耗,  ");
            sbSql.AppendLine("P19废铁单耗, P20回收率, P21熟料比, P22平均风温, P23炉顶温度, P24热风压力, ");
            sbSql.AppendLine("P25炉顶压力, P26富氧率, P27入炉焦炭总耗, P28入炉焦炭单耗, P29煤粉总耗, ");
            sbSql.AppendLine("P30煤粉单耗, P31焦丁总耗, P32焦丁单耗, P33综合焦炭总耗, P34综合焦炭单耗, ");
            sbSql.AppendLine("P35综合折算焦比, P36冶炼强度, P37焦炭负荷, P38干毛焦, P39炼钢铁SI, ");
            sbSql.AppendLine("P40炼钢铁MN, P41炼钢铁P, P42炼钢铁S, P43铸造铁SI, P44铸造铁MN, P45铸造铁P, ");
            sbSql.AppendLine("P46铸造铁S, P47炉渣碱度, P48休风情况, P49慢风, P50坐料次数, P51悬料次数, ");
            sbSql.AppendLine("P52崩料次数, P53风口损坏数大, P54风口损坏数中, P55风口损坏数小,  ");
            sbSql.AppendLine("P56渣口损坏数大, P57渣口损坏数中, P58渣口损坏数小, P59本厂铸造SI大, ");
            sbSql.AppendLine("P60本厂铸造SI小, P61送炼钢厂SI大, P62送炼钢厂SI中, P63送炼钢厂SI小,  ");
            sbSql.AppendLine("P64折算产量, P65工艺称焦比, P生成标志, P66S小于002, P67P小于009,  P68TI小于055,");
            sbSql.AppendLine("P69PB块, P70纽曼块, P71钛球, P72锰矿, P73硅石, P74白云石,P75蛇纹石,P76萤石,P77球团矿,P78国内球团矿,P79进口球团矿,");
            sbSql.AppendLine("P80其它块矿,P81高钛球团矿,P82高品位钛球,P83其它熔剂");
            sbSql.AppendLine("FROM 技术日报 WHERE (P日期 = :RQ)");

        //     P76萤石;

        //internal double? P77球团矿;

        //internal double? P78国内球团矿;

        //internal double? P79进口球团矿;
        //internal double? P80其它块矿;
        //internal double? P81高钛球团矿;

        //internal double? P82高品位钛球;
        //internal double? P83其它熔剂;

            conn.Open();
            OracleCommand cmdQuery = new OracleCommand(sbSql.ToString(), conn);
            cmdQuery.Parameters.Add(":RQ", OracleType.DateTime).Value = riqi;
            OracleDataReader dr = cmdQuery.ExecuteReader();
            while (dr.Read())
            {
                技术日报内容项 x = new 技术日报内容项(dr.GetString(0), dr.GetString(1));

                x.P03合格铁 = dr.IsDBNull(2) ? null : (double?)dr.GetDouble(2);
                x.P04炼钢铁 = dr.IsDBNull(3) ? null : (double?)dr.GetDouble(3);
                x.P05铸造铁 = dr.IsDBNull(4) ? null : (double?)dr.GetDouble(4);
                x.P06号外铁 = dr.IsDBNull(5) ? null : (double?)dr.GetDouble(5);
                x.P07合格率 = dr.IsDBNull(6) ? null : (double?)dr.GetDouble(6);
                x.P08高炉利用系数 = dr.IsDBNull(7) ? null : (double?)dr.GetDouble(7);
                x.P09高炉实物系数 = dr.IsDBNull(8) ? null : (double?)dr.GetDouble(8);
                x.P10原料矿合计总耗 = dr.IsDBNull(9) ? null : (double?)dr.GetDouble(9);
                x.P11原料矿合计单耗 = dr.IsDBNull(10) ? null : (double?)dr.GetDouble(10);
                x.P12原料矿机烧 = dr.IsDBNull(11) ? null : (double?)dr.GetDouble(11);
                x.P13原料矿竖炉球 = dr.IsDBNull(12) ? null : (double?)dr.GetDouble(12);
                x.P14原料矿CT = dr.IsDBNull(13) ? null : (double?)dr.GetDouble(13);
                x.P15原料矿其它熟料 = dr.IsDBNull(14) ? null : (double?)dr.GetDouble(14);
                x.P16原料矿本溪矿 = dr.IsDBNull(15) ? null : (double?)dr.GetDouble(15);
                x.P17原料矿其它生料 = dr.IsDBNull(16) ? null : (double?)dr.GetDouble(16);
                x.P18废铁总耗 = dr.IsDBNull(17) ? null : (double?)dr.GetDouble(17);
                x.P19废铁单耗 = dr.IsDBNull(18) ? null : (double?)dr.GetDouble(18);
                x.P20回收率 = dr.IsDBNull(19) ? null : (double?)dr.GetDouble(19);
                x.P21熟料比 = dr.IsDBNull(20) ? null : (double?)dr.GetDouble(20);
                x.P22平均风温 = dr.IsDBNull(21) ? null : (double?)dr.GetDouble(21);
                x.P23炉顶温度 = dr.IsDBNull(22) ? null : (double?)dr.GetDouble(22);
                x.P24热风压力 = dr.IsDBNull(23) ? null : (double?)dr.GetDouble(23);
                x.P25炉顶压力 = dr.IsDBNull(24) ? null : (double?)dr.GetDouble(24);
                x.P26富氧率 = dr.IsDBNull(25) ? null : (double?)dr.GetDouble(25);
                x.P27入炉焦炭总耗 = dr.IsDBNull(26) ? (int) 0 : (long)dr.GetDouble(26);
                x.P28入炉焦炭单耗 = dr.IsDBNull(27) ? null : (double?)dr.GetDouble(27);
                x.P29煤粉总耗 = dr.IsDBNull(28) ? (int)0 : (long)dr.GetDouble(28);
                x.P30煤粉单耗 = dr.IsDBNull(29) ? null : (double?)dr.GetDouble(29);
                x.P31焦丁总耗 = dr.IsDBNull(30) ? null : (double?)dr.GetDouble(30);
                x.P32焦丁单耗 = dr.IsDBNull(31) ? null : (double?)dr.GetDouble(31);
                x.P33综合焦炭总耗 = dr.IsDBNull(32) ? null : (double?)dr.GetDouble(32);
                x.P34综合焦炭单耗 = dr.IsDBNull(33) ? null : (double?)dr.GetDouble(33);
                x.P35综合折算焦比 = dr.IsDBNull(34) ? null : (double?)dr.GetDouble(34);
                x.P36冶炼强度 = dr.IsDBNull(35) ? null : (double?)dr.GetDouble(35);
                x.P37焦炭负荷 = dr.IsDBNull(36) ? null : (double?)dr.GetDouble(36);
                x.P38干毛焦 = dr.IsDBNull(37) ? (int)0 : (long)dr.GetDouble(37);
                x.P39炼钢铁SI = dr.IsDBNull(38) ? null : (double?)dr.GetDouble(38);
                x.P40炼钢铁MN = dr.IsDBNull(39) ? null : (double?)dr.GetDouble(39);
                x.P41炼钢铁P = dr.IsDBNull(40) ? null : (double?)dr.GetDouble(40);
                x.P42炼钢铁S = dr.IsDBNull(41) ? null : (double?)dr.GetDouble(41);
                x.P43铸造铁SI = dr.IsDBNull(42) ? null : (double?)dr.GetDouble(42);
                x.P44铸造铁MN = dr.IsDBNull(43) ? null : (double?)dr.GetDouble(43);
                x.P45铸造铁P = dr.IsDBNull(44) ? null : (double?)dr.GetDouble(44);
                x.P46铸造铁S = dr.IsDBNull(45) ? null : (double?)dr.GetDouble(45);
                x.P47炉渣碱度 = dr.IsDBNull(46) ? null : (double?)dr.GetDouble(46);
                x.P48休风情况 = dr.IsDBNull(47) ? null : (double?)dr.GetDouble(47);
                x.P49慢风 = dr.IsDBNull(48) ? null : (double?)dr.GetDouble(48);
                x.P50坐料次数 = dr.IsDBNull(49) ? null : (double?)dr.GetDouble(49);
                x.P51悬料次数 = dr.IsDBNull(50) ? null : (double?)dr.GetDouble(50);
                x.P52崩料次数 = dr.IsDBNull(51) ? null : (double?)dr.GetDouble(51);
                x.P53风口损坏数大 = dr.IsDBNull(52) ? null : (double?)dr.GetDouble(52);
                x.P54风口损坏数中 = dr.IsDBNull(53) ? null : (double?)dr.GetDouble(53);
                x.P55风口损坏数小 = dr.IsDBNull(54) ? null : (double?)dr.GetDouble(54);
                x.P56渣口损坏数大 = dr.IsDBNull(55) ? null : (double?)dr.GetDouble(55);
                x.P57渣口损坏数中 = dr.IsDBNull(56) ? null : (double?)dr.GetDouble(56);
                x.P58渣口损坏数小 = dr.IsDBNull(57) ? null : (double?)dr.GetDouble(57);
                x.P59本厂铸造SI大 = dr.IsDBNull(58) ? null : (double?)dr.GetDouble(58);
                x.P60本厂铸造SI小 = dr.IsDBNull(59) ? null : (double?)dr.GetDouble(59);
                x.P61送炼钢厂SI大 = dr.IsDBNull(60) ? null : (double?)dr.GetDouble(60);
                x.P62送炼钢厂SI中 = dr.IsDBNull(61) ? null : (double?)dr.GetDouble(61);
                x.P63送炼钢厂SI小 = dr.IsDBNull(62) ? null : (double?)dr.GetDouble(62);
                x.P64折算产量 = dr.IsDBNull(63) ? null : (double?)dr.GetDouble(63);
                x.P65工艺称焦比 = dr.IsDBNull(64) ? null : (double?)dr.GetDouble(64);
                x.P生成标志 = dr.IsDBNull(65) ? null : (double?)dr.GetDouble(65);
                x.P66S小于002 = dr.IsDBNull(66) ? null : (double?)dr.GetDouble(66);
                x.P67P小于009 = dr.IsDBNull(67) ? null : (double?)dr.GetDouble(67);
                x.P68TI小于055 = dr.IsDBNull(68) ? null : (double?)dr.GetDouble(68);
                //x.P69PB块 = dr.IsDBNull(69) ? null : (double?)dr.GetDouble(69);
                //x.P70纽曼块 = dr.IsDBNull(70) ? null : (double?)dr.GetDouble(70);
                //x.P71钛球= dr.IsDBNull(71) ? null : (double?)dr.GetDouble(71);
                //x.P72锰矿 = dr.IsDBNull(72) ? null : (double?)dr.GetDouble(72);
                //x.P73硅石 = dr.IsDBNull(73) ? null : (double?)dr.GetDouble(73);
                //x.P74白云石 = dr.IsDBNull(74) ? null : (double?)dr.GetDouble(74);
                //x.P75蛇纹石 = dr.IsDBNull(75) ? null : (double?)dr.GetDouble(75);

                //x.P76萤石 = dr.IsDBNull(76) ? null : (double?)dr.GetDouble(76);
                //x.P77球团矿 = dr.IsDBNull(77) ? null : (double?)dr.GetDouble(77);
                //x.P78国内球团矿 = dr.IsDBNull(78) ? null : (double?)dr.GetDouble(78);
                //x.P79进口球团矿 = dr.IsDBNull(79) ? null : (double?)dr.GetDouble(79);
                //x.P80其它块矿 = dr.IsDBNull(80) ? null : (double?)dr.GetDouble(80);
                //x.P81高钛球团矿 = dr.IsDBNull(81) ? null : (double?)dr.GetDouble(81);
                //x.P82高品位钛球 = dr.IsDBNull(82) ? null : (double?)dr.GetDouble(82);
                //x.P83其它熔剂 = dr.IsDBNull(83) ? null : (double?)dr.GetDouble(83);

                x.P69PB块 = dr.IsDBNull(69) ? (int)0 : (long)dr.GetDouble(69);
                x.P70纽曼块 = dr.IsDBNull(70) ? (int)0 : (long)dr.GetDouble(70);
                x.P71钛球 = dr.IsDBNull(71) ? (int)0 : (long)dr.GetDouble(71);
                x.P72锰矿 = dr.IsDBNull(72) ? (int)0 : (long)dr.GetDouble(72);
                x.P73硅石 = dr.IsDBNull(73) ? (int)0 : (long)dr.GetDouble(73);
                x.P74白云石 = dr.IsDBNull(74) ? (int)0 : (long)dr.GetDouble(74);
                x.P75蛇纹石 = dr.IsDBNull(75) ? (int)0 : (long)dr.GetDouble(75);
                x.P76萤石 = dr.IsDBNull(76) ? null : (double?)dr.GetDouble(76);
                x.P77球团矿 = dr.IsDBNull(77) ? (int)0 : (long)dr.GetDouble(77);
                x.P78国内球团矿 = dr.IsDBNull(78) ? (int)0 : (long)dr.GetDouble(78);
                x.P79进口球团矿 = dr.IsDBNull(79) ? (int)0 : (long)dr.GetDouble(79);
                x.P80其它块矿 = dr.IsDBNull(80) ? (int)0 : (long)dr.GetDouble(80);
                x.P81高钛球团矿 = dr.IsDBNull(81) ? (int)0 : (long)dr.GetDouble(81);
                x.P82高品位钛球 = dr.IsDBNull(82) ? (int)0 : (long)dr.GetDouble(82);
                x.P83其它熔剂 = dr.IsDBNull(83) ? (int)0 : (long)dr.GetDouble(83);

                result.Add(x);
            }
            //     P76萤石;

            //internal double? P77球团矿;

            //internal double? P78国内球团矿;

            //internal double? P79进口球团矿;
            //internal double? P80其它块矿;
            //internal double? P81高钛球团矿;

            //internal double? P82高品位钛球;
            //internal double? P83其它熔剂;
            dr.Close();
            conn.Close();
            return result;
        }

        public void Save(技术日报内容 n)
        {
            conn.Open();
            

            conn.Close();
      
        }
    }
}
