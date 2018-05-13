using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;

namespace LTZN.技术日报
{
    class 生铁
    {
        public double[] 炼钢铁 = new double[6];
        public double[] 铸造铁 = new double[6];
        public double[] 合格铁 = new double[6];
        public double[] 号外铁= new double[6];
        public double[] 合格率= new double[6];
        public double[] 高炉利用系数= new double[6];
        public double[] 高炉实物系数= new double[6];
        public double[] 炼钢铁Si= new double[6];
        public double[] 炼钢铁Mn = new double[6];
        public double[] 炼钢铁P = new double[6];
        public double[] 炼钢铁S= new double[6];
        public double[] 铸造铁Si = new double[6];
        public double[] 铸造铁Mn= new double[6];
        public double[] 铸造铁P = new double[6];
        public double[] 铸造铁S= new double[6];
        public double[] 炉渣碱度= new double[6];
        public double[] 本厂铸造SI大 = new double[6];
        public double[] 本厂铸造SI小= new double[6];
        public double[] 送炼钢厂SI大 = new double[6];
        public double[] 送炼钢厂SI小= new double[6];
        public double[] 折算产量 = new double[6];

        public double[] 累计炼钢铁 = new double[6];
        public double[] 累计铸造铁 = new double[6];
        public double[] 累计合格铁 = new double[6];
        public double[] 累计号外铁 = new double[6];
        public double[] 累计合格率 = new double[6];
        public double[] 累计高炉利用系数 = new double[6];
        public double[] 累计高炉实物系数 = new double[6];
        public double[] 累计炼钢铁Si = new double[6];
        public double[] 累计炼钢铁Mn = new double[6];
        public double[] 累计炼钢铁P = new double[6];
        public double[] 累计炼钢铁S = new double[6];
        public double[] 累计铸造铁Si = new double[6];
        public double[] 累计铸造铁Mn = new double[6];
        public double[] 累计铸造铁P = new double[6];
        public double[] 累计铸造铁S = new double[6];
        public double[] 累计炉渣碱度 = new double[6];
        public double[] 累计本厂铸造SI大 = new double[6];
        public double[] 累计本厂铸造SI小 = new double[6];
        public double[] 累计送炼钢厂SI大 = new double[6];
        public double[] 累计送炼钢厂SI小 = new double[6];
        public double[] 累计折算产量 = new double[6];

        public double[] 全铁产量 = new double[6];
        public double[] 累计全铁产量 = new double[6];
        public double[] 累计大中修= new double[6];
        public double[] 大中修 = new double[6];
        public double[] 累计有效炉容 = new double[6];
        public double[] 有效炉容 = new double[6];
        public void getDataBy(DateTime rq)
        {
            OracleConnection cn = new OracleConnection();
            cn.ConnectionString = Properties.Settings.Default.ConnectionString;
            cn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = cn;


            //炼钢铁
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where fesi<=1.25 and trunc(zdsj)=:rq and dksj is not null group by gaolu";
            cmd.Parameters.Add(":rq", OracleType.DateTime).Value = rq;
            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    炼钢铁[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
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

            //铸造铁
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where fesi>1.25 and trunc(zdsj)=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    铸造铁[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                }
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
                    合格铁[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                }
            }
            dr.Close();

            //送炼钢厂SI小
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where fesi<=1.25 and quchu='炼钢' and trunc(zdsj)=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    送炼钢厂SI小[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                }
            }
            dr.Close();

            //送炼钢厂SI大
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where fesi>1.25 and quchu='炼钢' and trunc(zdsj)=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    送炼钢厂SI大[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                }
            }
            dr.Close();

            //本厂铸造SI小
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where fesi<=1.25 and quchu='铸铁' and trunc(zdsj)=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    本厂铸造SI小[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                }
            }
            dr.Close();

            //本厂铸造SI大
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where fesi>1.25 and quchu='铸铁' and trunc(zdsj)=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    本厂铸造SI大[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                }
            }
            dr.Close();

            //炼钢铁成分
            cmd.CommandText = "select gaolu,trunc(avg(FESI),2),trunc(avg(FEMn),2),trunc(avg(FEP),3),trunc(avg(FES),3) from ddluci where fesi<=1.25 and trunc(zdsj)=:rq and dksj is not null group by gaolu";
            cmd.Parameters.Add(":rq", OracleType.DateTime).Value = rq;
             dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    炼钢铁Si[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    炼钢铁Mn[gaolu - 1] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                    炼钢铁P[gaolu - 1] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                    炼钢铁S[gaolu - 1] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                }
            }
            dr.Close();

            //铸造铁成分
            cmd.CommandText = "select gaolu,trunc(avg(FESI),2),trunc(avg(FEMn),2),trunc(avg(FEP),3),trunc(avg(FES),3) from ddluci where fesi>1.25 and trunc(zdsj)=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    铸造铁Si[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    铸造铁Mn[gaolu - 1] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                    铸造铁P[gaolu - 1] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                    铸造铁S[gaolu - 1] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                }
            }
            dr.Close();

            //全厂炼钢铁成分
            cmd.CommandText = "select trunc(avg(FESI),2),trunc(avg(FEMn),2),trunc(avg(FEP),3),trunc(avg(FES),3) from ddluci where fesi<=1.25 and trunc(zdsj)=:rq and dksj is not null";
            cmd.Parameters.Add(":rq", OracleType.DateTime).Value = rq;
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    炼钢铁Si[5] = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
                    炼钢铁Mn[5] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    炼钢铁P[5] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                    炼钢铁S[5] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                }
            }
            dr.Close();

            //全厂铸造铁成分
            cmd.CommandText = "select trunc(avg(FESI),2),trunc(avg(FEMn),2),trunc(avg(FEP),3),trunc(avg(FES),3) from ddluci where fesi>1.25 and trunc(zdsj)=:rq and dksj is not null";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    铸造铁Si[5] = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
                    铸造铁Mn[5] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    铸造铁P[5] = dr.IsDBNull(2) ? 0 : dr.GetDouble(1);
                    铸造铁S[5] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                }
            }
            dr.Close();


            //炉渣成分
            cmd.CommandText = "select gaolu,trunc(avg(ZHAR2),2) from ddluci where trunc(zdsj)=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    炉渣碱度[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                   
                }
            }
            dr.Close();

            //全厂炉渣成分
            cmd.CommandText = "select trunc(avg(ZHAR2),2) from ddluci where  trunc(zdsj)=:rq and dksj is not null";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    炉渣碱度[5] = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
                }
            }
            dr.Close();

            //全铁产量
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where  trunc(zdsj)=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    全铁产量[gaolu-1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
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
                    折算产量[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                }
            }
            dr.Close();

            for (int i = 0; i < 5; i++)
            {
                全铁产量[5] += 全铁产量[i];
                炼钢铁[5] += 炼钢铁[i];
                铸造铁[5] += 铸造铁[i];
                合格铁[5] += 合格铁[i];
                号外铁[i] = 全铁产量[i] - 合格铁[i];
                号外铁[5] += 号外铁[i];
                送炼钢厂SI大[5] += 送炼钢厂SI大[i];
                送炼钢厂SI小[5] += 送炼钢厂SI小[i];
                本厂铸造SI大[5] += 本厂铸造SI大[i];
                本厂铸造SI小[5] += 本厂铸造SI小[i];
                折算产量[5] += 折算产量[i];
            }
            有效炉容[0] = (((1440 * 1) - 大中修[0]) / (1440 * 1)) * Constants.volume1;
            有效炉容[1] = (((1440 * 1) - 大中修[1]) / (1440 * 1)) * Constants.volume2;
            有效炉容[2] = (((1440 * 1) - 大中修[2]) / (1440 * 1)) * Constants.volume3;
            有效炉容[3] = (((1440 * 1) - 大中修[3]) / (1440 * 1)) * Constants.volume4;
            有效炉容[4] = (((1440 * 1) - 大中修[4]) / (1440 * 1)) * Constants.volume5;
            有效炉容[5] = 有效炉容[0] + 有效炉容[1] + 有效炉容[2] + 有效炉容[3] + 有效炉容[4];

            高炉利用系数[0] = 折算产量[0] / 有效炉容[0];
            高炉利用系数[1] = 折算产量[1] / 有效炉容[1];
            高炉利用系数[2] = 折算产量[2] / 有效炉容[2];
            高炉利用系数[3] = 折算产量[3] / 有效炉容[3];
            高炉利用系数[4] = 折算产量[4] / 有效炉容[4];
            高炉利用系数[5] = 折算产量[5] / 有效炉容[5];

            高炉实物系数[0] = 合格铁[0] / 有效炉容[0];
            高炉实物系数[1] = 合格铁[1] / 有效炉容[1];
            高炉实物系数[2] = 合格铁[2] / 有效炉容[2];
            高炉实物系数[3] = 合格铁[3] / 有效炉容[3];
            高炉实物系数[4] = 合格铁[4] / 有效炉容[4];
            高炉实物系数[5] = 合格铁[5] / 有效炉容[5];
            for (int i = 0; i < 6; i++)
            {
                if (全铁产量[i]>0)
                  合格率[i] = 合格铁[i]*100 / 全铁产量[i];
            }

            //累计部分
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
            //累计炼钢铁
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where fesi<=1.25 and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    累计炼钢铁[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                }
            }
            dr.Close();

            //累计铸造铁
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where fesi>1.25 and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    累计铸造铁[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                }
            }
            dr.Close();

            //累计合格铁
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where  合格铁(fesi,fes)=1 and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    累计合格铁[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                }
            }
            dr.Close();

            //累计送炼钢厂SI小
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where fesi<=1.25 and quchu='炼钢' and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    累计送炼钢厂SI小[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                }
            }
            dr.Close();

            //累计送炼钢厂SI大
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where fesi>1.25 and quchu='炼钢' and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    累计送炼钢厂SI大[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                }
            }
            dr.Close();

            //累计本厂铸造SI小
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where fesi<=1.25 and quchu='铸铁' and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    累计本厂铸造SI小[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                }
            }
            dr.Close();

            //累计本厂铸造SI大
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where fesi>1.25 and quchu='铸铁' and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    累计本厂铸造SI大[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                }
            }
            dr.Close();

            //累计炼钢铁成分
            cmd.CommandText = "select gaolu,trunc(avg(FESI),2),trunc(avg(FEMn),2),trunc(avg(FEP),3),trunc(avg(FES),3) from ddluci where fesi<=1.25 and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and dksj is not null group by gaolu";
            cmd.Parameters.Add(":rq", OracleType.DateTime).Value = rq;
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    累计炼钢铁Si[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    累计炼钢铁Mn[gaolu - 1] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                    累计炼钢铁P[gaolu - 1] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                    累计炼钢铁S[gaolu - 1] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                }
            }
            dr.Close();

            //累计铸造铁成分
            cmd.CommandText = "select gaolu,trunc(avg(FESI),2),trunc(avg(FEMn),2),trunc(avg(FEP),3),trunc(avg(FES),3) from ddluci where fesi>1.25 and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    累计铸造铁Si[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    累计铸造铁Mn[gaolu - 1] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                    累计铸造铁P[gaolu - 1] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                    累计铸造铁S[gaolu - 1] = dr.IsDBNull(4) ? 0 : dr.GetDouble(4);
                }
            }
            dr.Close();

            //累计全厂炼钢铁成分
            cmd.CommandText = "select trunc(avg(FESI),2),trunc(avg(FEMn),2),trunc(avg(FEP),3),trunc(avg(FES),3) from ddluci where fesi<=1.25 and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and dksj is not null";
            cmd.Parameters.Add(":rq", OracleType.DateTime).Value = rq;
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    累计炼钢铁Si[5] = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
                    累计炼钢铁Mn[5] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    累计炼钢铁P[5] = dr.IsDBNull(2) ? 0 : dr.GetDouble(2);
                    累计炼钢铁S[5] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                }
            }
            dr.Close();

            //累计全厂铸造铁成分
            cmd.CommandText = "select trunc(avg(FESI),2),trunc(avg(FEMn),2),trunc(avg(FEP),3),trunc(avg(FES),3) from ddluci where fesi>1.25 and trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and dksj is not null";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    累计铸造铁Si[5] = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
                    累计铸造铁Mn[5] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                    累计铸造铁P[5] = dr.IsDBNull(2) ? 0 : dr.GetDouble(1);
                    累计铸造铁S[5] = dr.IsDBNull(3) ? 0 : dr.GetDouble(3);
                }
            }
            dr.Close();


            //累计炉渣成分
            cmd.CommandText = "select gaolu,trunc(avg(ZHAR2),2) from ddluci where trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    累计炉渣碱度[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);

                }
            }
            dr.Close();

            //累计全厂炉渣成分
            cmd.CommandText = "select trunc(avg(ZHAR2),2) from ddluci where  trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and dksj is not null";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    累计炉渣碱度[5] = dr.IsDBNull(0) ? 0 : dr.GetDouble(0);
                }
            }
            dr.Close();

            //累计全铁产量
            cmd.CommandText = "select gaolu,sum(feliang) from ddluci where  trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    累计全铁产量[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                }
            }
            dr.Close();

            
            //累计折算产量
            cmd.CommandText = "select gaolu,sum(折算产量(feliang,fesi,fes)) from ddluci where trunc(zdsj,'MONTH')=trunc(:rq,'MONTH') and trunc(zdsj)<=:rq and dksj is not null group by gaolu";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    int gaolu = dr.GetInt32(0);
                    累计折算产量[gaolu - 1] = dr.IsDBNull(1) ? 0 : dr.GetDouble(1);
                }
            }
            dr.Close();

            for (int i = 0; i < 5; i++)
            {
                累计全铁产量[5] += 累计全铁产量[i];
                累计炼钢铁[5] += 累计炼钢铁[i];
                累计铸造铁[5] += 累计铸造铁[i];
                累计合格铁[5] += 累计合格铁[i];
                累计号外铁[i] = 累计全铁产量[i] - 累计合格铁[i];
                累计号外铁[5] += 累计号外铁[i];
                累计送炼钢厂SI大[5] += 累计送炼钢厂SI大[i];
                累计送炼钢厂SI小[5] += 累计送炼钢厂SI小[i];
                累计本厂铸造SI大[5] += 累计本厂铸造SI大[i];
                累计本厂铸造SI小[5] += 累计本厂铸造SI小[i];
                累计折算产量[5] += 累计折算产量[i];
            }
            累计有效炉容[0] = (((1440 * rq.Day) - 累计大中修[0]) / (1440 * rq.Day)) * Constants.volume1 * rq.Day;
            累计有效炉容[1] = (((1440 * rq.Day) - 累计大中修[1]) / (1440 * rq.Day)) * Constants.volume2 * rq.Day;
            累计有效炉容[2] = (((1440 * rq.Day) - 累计大中修[2]) / (1440 * rq.Day)) * Constants.volume3 * rq.Day;
            累计有效炉容[3] = (((1440 * rq.Day) - 累计大中修[3]) / (1440 * rq.Day)) * Constants.volume4 * rq.Day;
            累计有效炉容[4] = (((1440 * rq.Day) - 累计大中修[4]) / (1440 * rq.Day)) * Constants.volume5 * rq.Day;
            累计有效炉容[5] = 累计有效炉容[0] + 累计有效炉容[1] + 累计有效炉容[2] + 累计有效炉容[3] + 累计有效炉容[4];

            累计高炉利用系数[0] = 累计折算产量[0] / 累计有效炉容[0];
            累计高炉利用系数[1] = 累计折算产量[1] / 累计有效炉容[1];
            累计高炉利用系数[2] = 累计折算产量[2] / 累计有效炉容[2];
            累计高炉利用系数[3] = 累计折算产量[3] / 累计有效炉容[3];
            累计高炉利用系数[4] = 累计折算产量[4] / 累计有效炉容[4];
            累计高炉利用系数[5] = 累计折算产量[5] / 累计有效炉容[5];

            累计高炉实物系数[0] = 累计合格铁[0] / 累计有效炉容[0];
            累计高炉实物系数[1] = 累计合格铁[1] / 累计有效炉容[1];
            累计高炉实物系数[2] = 累计合格铁[2] / 累计有效炉容[2];
            累计高炉实物系数[3] = 累计合格铁[3] / 累计有效炉容[3];
            累计高炉实物系数[4] = 累计合格铁[4] / 累计有效炉容[4];
            累计高炉实物系数[5] = 累计合格铁[5] / 累计有效炉容[5];
            for (int i = 0; i < 6; i++)
            {
                if (累计全铁产量[i] > 0)
                    累计合格率[i] = 累计合格铁[i]*100 / 累计全铁产量[i];
            }
            cn.Close();
        }

    }
}
