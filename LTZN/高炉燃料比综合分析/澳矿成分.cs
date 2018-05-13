using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;

namespace LTZN.高炉燃料比综合分析
{
    public partial class 澳矿成分录入 : Form
    {

        bool isdata = false;    //判断是否有澳矿成分
        bool ztisdata = false;    //判断是否有渣铁成分
        bool stisdata = false;    //判断是否有渣铁成分
        bool yhtpisdata = false;    //判断是否有渣铁成分
        bool tqisdata = false;    //判断是否有渣铁成分
        bool isP = false;       //判断是否有P
        OracleConnection con = new OracleConnection(Properties.Settings.Default.ltznConnectionString);

        public 澳矿成分录入()
        {
            InitializeComponent();
        }

       
        private void button1_Click(object sender, EventArgs e)
        {
            
            
       
        }

        private void 澳矿成分录入_Load(object sender, EventArgs e)
        {
            con.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = con;
            cmd.CommandText = "select * from ddyl where mc='澳矿'";
            OracleDataReader odr =cmd.ExecuteReader();
            while (odr.Read())
            {

                if (!(odr[1] is DBNull))
                {
                    isdata  = true;    //有澳矿

                }
                if (!(odr[4] is DBNull))
                {
                    this.textBox1.Text=odr[4].ToString();    //TFE

                }
                if (!(odr[5] is DBNull))
                {
                    this.textBox4.Text = odr[5].ToString();    //SIO

                }
                if (!(odr[6] is DBNull))
                {
                    this.textBox5.Text = odr[6].ToString();    //CAO

                }
                if (!(odr[7] is DBNull))
                {
                    this.textBox6.Text = odr[7].ToString();    //FEO

                }
                if (!(odr[8] is DBNull))
                {
                    this.textBox10.Text = odr[8].ToString();    //MGO

                }
                if (!(odr[9] is DBNull))
                {
                    this.textBox2.Text = odr[9].ToString();    //S

                }
                if (!(odr[10] is DBNull))
                {
                    this.textBox7.Text = odr[10].ToString();    //ALO

                }
                if (!(odr[11] is DBNull))
                {
                    this.textBox9.Text = odr[11].ToString();    //JD

                }
                if (!(odr[14] is DBNull))
                {
                    this.textBox8.Text = odr[14].ToString();    //TIO2

                }
                if (!(odr[15] is DBNull))
                {
                    this.textBox3.Text = odr[15].ToString();    //P

                }
            }
            odr.Close();
            cmd.CommandText = "select * from ddjtsetp where mc='焦炭'";
             odr =cmd.ExecuteReader();
            while (odr.Read())
            {
                if (!(odr[1] is DBNull))
                {
                    isP = true;    //有P

                }
                if (!(odr[2] is DBNull))
                {
                    this.textBox11.Text = odr[2].ToString();    //TFE

                }
            }
            odr.Close();
            cmd.CommandText = "select * from ddyl where mc='碎铁'";
            odr = cmd.ExecuteReader();
            while (odr.Read())
            {
                if (!(odr[4] is DBNull))
                {
                    stisdata = true;    //有碎铁

                }
                else
                {
                    stisdata = false;    //有碎铁
                }
                if (!(odr[4] is DBNull))
                {
                    this.textBox12.Text = odr[4].ToString();    //TFE

                }
            }
            odr.Close();
            cmd.CommandText = "select * from ddyl where mc='渣铁'";
            odr = cmd.ExecuteReader();
            while (odr.Read())
            {
                if (!(odr[4] is DBNull))
                {
                    ztisdata = true;    //有渣铁

                }
                else
                {
                    ztisdata = false;    //有渣铁
                }
                if (!(odr[4] is DBNull))
                {
                    this.textBox13.Text = odr[4].ToString();    //TFE

                }
            }
            cmd.CommandText = "select * from ddyl where mc='氧化铁皮'";
            odr = cmd.ExecuteReader();
            while (odr.Read())
            {
                if (!(odr[4] is DBNull))
                {
                    yhtpisdata = true;    //有氧化铁皮

                }
                else
                {
                    yhtpisdata = false;    //有渣铁
                }
                if (!(odr[4] is DBNull))
                {
                    this.textBox14.Text = odr[4].ToString();    //TFE

                }
            }
            cmd.CommandText = "select * from ddyl where mc='钛球'";
            odr = cmd.ExecuteReader();
            while (odr.Read())
            {
                if (!(odr[4] is DBNull))
                {
                    tqisdata = true;    //有钛球

                }
                else
                {
                    tqisdata = false;    //钛球
                }
                if (!(odr[4] is DBNull))
                {
                    this.textBox15.Text = odr[4].ToString();    //TFE

                }
            }
            odr.Close();
            con.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            con.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = con;
            if (isP == true)
            {
                cmd.CommandText = "update ddjtsetp set mc='焦炭',sj=sysdate,p=:P where mc='焦炭'";
                cmd.Parameters.Add(":P", OracleType.Double).Value = Convert.ToDouble(this.textBox11.Text);
              
              
                cmd.ExecuteNonQuery();
                MessageBox.Show("更新成功");
            }
            else
            {
                cmd.CommandText = "insert into ddjtsetp(mc,sj,p)" +
                                          " values ('焦炭',sysdate,:P)";
                cmd.Parameters.Add(":P", OracleType.Double).Value = Convert.ToDouble(this.textBox11.Text);
         
                cmd.ExecuteNonQuery();
                MessageBox.Show("插入完成");
            }
            con.Close();
        }
        public void saveaokuang()
        {
            con.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = con;
            if (isdata == true)
            {
                cmd.CommandText = "update ddyl set mc='澳矿',sj=sysdate,cang='1-4#',tfe=:TFE,sio=:SiO,cao=:CaO,feo=:FeO,mgo=:MgO,s=:S,alo=:Al203,jd=:JD,tio2=:Ti02,p=:P where mc='澳矿'";
                cmd.Parameters.Add(":TFE", OracleType.Double).Value = Convert.ToDouble(this.textBox1.Text);
                cmd.Parameters.Add(":SiO", OracleType.Double).Value = Convert.ToDouble(this.textBox4.Text);
                cmd.Parameters.Add("CaO", OracleType.Double).Value = Convert.ToDouble(this.textBox5.Text);
                cmd.Parameters.Add("FeO", OracleType.Double).Value = Convert.ToDouble(this.textBox6.Text);
                cmd.Parameters.Add(":MgO", OracleType.Double).Value = Convert.ToDouble(this.textBox10.Text);
                cmd.Parameters.Add(":S", OracleType.Double).Value = Convert.ToDouble(this.textBox2.Text);
                cmd.Parameters.Add(":Al203", OracleType.Double).Value = Convert.ToDouble(this.textBox7.Text);
                cmd.Parameters.Add(":JD", OracleType.Double).Value = Convert.ToDouble(this.textBox9.Text);
                cmd.Parameters.Add(":Ti02", OracleType.Double).Value = Convert.ToDouble(this.textBox8.Text);
                cmd.Parameters.Add(":P", OracleType.Double).Value = Convert.ToDouble(this.textBox3.Text);
                cmd.ExecuteNonQuery();


            }
            else
            {
                cmd.CommandText = "insert into ddyl(mc,sj,cang,tfe,sio,cao,feo,mgo,s,alo,jd,tio2,p)" +
                                          " values ('澳矿',sysdate,'1-4#',:TFE,:SiO,:CaO,:FeO,:MgO,:S,:Al203,:JD,:Ti02,:P)";
                cmd.Parameters.Add(":TFE", OracleType.Double).Value = Convert.ToDouble(this.textBox1.Text);
                cmd.Parameters.Add(":SiO", OracleType.Double).Value = Convert.ToDouble(this.textBox4.Text);
                cmd.Parameters.Add("CaO", OracleType.Double).Value = Convert.ToDouble(this.textBox5.Text);
                cmd.Parameters.Add("FeO", OracleType.Double).Value = Convert.ToDouble(this.textBox6.Text);
                cmd.Parameters.Add(":MgO", OracleType.Double).Value = Convert.ToDouble(this.textBox10.Text);
                cmd.Parameters.Add(":S", OracleType.Double).Value = Convert.ToDouble(this.textBox2.Text);
                cmd.Parameters.Add(":Al203", OracleType.Double).Value = Convert.ToDouble(this.textBox7.Text);
                cmd.Parameters.Add(":JD", OracleType.Double).Value = Convert.ToDouble(this.textBox9.Text);
                cmd.Parameters.Add(":Ti02", OracleType.Double).Value = Convert.ToDouble(this.textBox8.Text);
                cmd.Parameters.Add(":P", OracleType.Double).Value = Convert.ToDouble(this.textBox3.Text);
                cmd.ExecuteNonQuery();


            }
            con.Close();
        }
        public void saveqita(Boolean istrue,string name,double value)
        {
            con.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = con;
            if (istrue == true)
           {
                //渣铁
                cmd.CommandText = " update ddyl set sj=sysdate,tfe=:TFE where mc=:mc";
                cmd.Parameters.Add(":TFE", OracleType.Double).Value = value;
                cmd.Parameters.Add(":mc", OracleType.NVarChar).Value =name;
                cmd.ExecuteNonQuery();
            }
            else
            {
                //渣铁
                cmd.CommandText = "insert into ddyl(mc,sj,cang,tfe,sio,cao,feo,mgo,s,alo,jd,tio2,p) values (:mc,sysdate,'1-4#',:TFE,0,0,0,0,0,0,0,0,0)";
                cmd.Parameters.Add(":TFE", OracleType.Double).Value = value;
                cmd.Parameters.Add(":mc", OracleType.NVarChar).Value =name;
                cmd.ExecuteNonQuery();
              
            }
            con.Close();
        }
        private void button1_Click_1(object sender, EventArgs e)
        {
            saveaokuang();

            saveqita(ztisdata,"渣铁",Convert.ToDouble(textBox13.Text));
            saveqita(stisdata, "碎铁", Convert.ToDouble(textBox12.Text));
            saveqita(yhtpisdata, "氧化铁皮", Convert.ToDouble(textBox14.Text));
            saveqita(tqisdata, "钛球", Convert.ToDouble(textBox15.Text));
         
        
            MessageBox.Show("操作完成");
        }

    }
}
