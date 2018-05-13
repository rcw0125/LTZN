using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data.OracleClient;

namespace LTZN.查询
{
    public partial class 数据查询窗口 : Form
    {
        private string tiaojian = "";
        
        public 数据查询窗口()
        {
            InitializeComponent();
            LtznUserManager.instance.RegisterHandler(this, UserChange);
        }

        //用户变动时
        void UserChange(LtznUser ltznUser)
        {
            if (ltznUser != null && ltznUser.IsInRole("统计"))
            {
                viewCalModel.Visible = true;
                viewCalModel.Enabled = true;
            }
            else
            {
                viewCalModel.Visible = false;
                viewCalModel.Enabled = false;
            }
        }

        private 本日生铁 _benRiShengTie = null;
        public 本日生铁 BenRiShengTie
        {
            get
            {
                if (_benRiShengTie == null)
                {
                    _benRiShengTie = new 本日生铁();
                }
                return _benRiShengTie;
            }
        }
        
        private 上报厂调 _shangBaoChangDiao = null;
        public 上报厂调 ShangBaoChangDiao
        {
            get
            {
                if (_shangBaoChangDiao == null)
                {
                    _shangBaoChangDiao = new 上报厂调();
                }
                return _shangBaoChangDiao;
            }
        }

        private 其它数据 _qiTaShuJu = null;
        public 其它数据 QiTaShuJu
        {
            get {
                if (_qiTaShuJu == null)
                {
                    _qiTaShuJu = new 其它数据();
                }
                return _qiTaShuJu;
            }
        }

        private LTZN.CalFramework.CalModel calModel = null;

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

            this.Cursor = Cursors.WaitCursor;

            //本日生铁
            if (this.radioButton1.Checked)
            {
                this.BenRiShengTie.getData(dateTimePicker1.Value.Date);
                this.BenRiShengTie.ModelCalc();
                this.BenRiShengTie.display(webBrowser1.Document);
            }

            //上报厂调
            if (this.radioButton2.Checked)
            {
               this.ShangBaoChangDiao.getData(dateTimePicker1.Value.Date);
               this.ShangBaoChangDiao.ModelCalc();
               this.ShangBaoChangDiao.display(webBrowser1.Document);
            }

            //其它数据
            if (this.radioButton3.Checked)
            {
                this.QiTaShuJu.getData(dateTimePicker1.Value.Date);
                this.QiTaShuJu.ModelCalc();
                this.QiTaShuJu.display(webBrowser1.Document);
            }
            if (this.radioButton4.Checked)
            {
                tieliangkaohe();
            }
            if (this.radioButton5.Checked)
            {
                优质优价炼钢铁 a = new 优质优价炼钢铁();
                a.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date, tiaojian);
                a.display(webBrowser1.Document);
            }
            if (this.radioButton6.Checked)
            {
                机烧值比较 a = new 机烧值比较();
                a.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date);
                a.display(webBrowser1.Document);
            }
            if (this.radioButton7.Checked)
            {
                优质铁水率 a = new 优质铁水率();
                a.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date);
                a.display(webBrowser1.Document);
            }
            if (this.radioButton8.Checked)
            {
                生产指标 a = new 生产指标();
                a.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date);
                a.display(webBrowser1.Document);
            }
            if (this.radioButton9.Checked)
            {
                晚点时间 a = new 晚点时间();
                a.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date);
                a.display(webBrowser1.Document);
            }
            if (this.radioButton10.Checked)
            {
                大宗原料 a = new 大宗原料();
                a.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date);
                a.display(webBrowser1.Document);
            }
            if (this.radioButton11.Checked)
            {
                铁口深度 a = new 铁口深度();
                a.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date);
                a.display(webBrowser1.Document);
            }
            if (this.radioButton12.Checked)
            {
                风温考核 a = new 风温考核();
                a.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date);
                a.display(webBrowser1.Document);
            }
            this.Cursor = Cursors.Default;
        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            if (this.radioButton4.Checked)
            {
                tieliangkaohe();
            }
            if (this.radioButton5.Checked)
            {
                优质优价炼钢铁 a = new 优质优价炼钢铁();
                a.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date, tiaojian);
                a.display(webBrowser1.Document);
            }
            if (this.radioButton13.Checked)
            {
              
                铁水质量查询 a = new 铁水质量查询();
                a.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date, tiaojian);
                a.display(webBrowser1.Document);
                
            }
            if (this.radioButton6.Checked)
            {
                机烧值比较 a = new 机烧值比较();
                a.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date);
                a.display(webBrowser1.Document);
            }
            if (this.radioButton7.Checked)
            {
                优质铁水率 a = new 优质铁水率();
                a.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date);
                a.display(webBrowser1.Document);
            }
            if (this.radioButton8.Checked)
            {
                生产指标 a = new 生产指标();
                a.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date);
                a.display(webBrowser1.Document);
            }
            if (this.radioButton9.Checked)
            {
                晚点时间 a = new 晚点时间();
                a.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date);
                a.display(webBrowser1.Document);
            }
            if (this.radioButton10.Checked)
            {
                大宗原料 a = new 大宗原料();
                a.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date);
                a.display(webBrowser1.Document);
            }
            if (this.radioButton11.Checked)
            {
                铁口深度 a = new 铁口深度();
                a.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date);
                a.display(webBrowser1.Document);
            }
            if (this.radioButton12.Checked)
            {
                风温考核 a = new 风温考核();
                a.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date);
                a.display(webBrowser1.Document);
            }
            this.Cursor = Cursors.Default;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            if (this.radioButton1.Checked)
            {
                dateTimePicker2.Visible = false;
                this.BenRiShengTie.getData(dateTimePicker1.Value.Date);
                this.BenRiShengTie.ModelCalc();
                this.BenRiShengTie.display(webBrowser1.Document);
                this.calModel = this.BenRiShengTie.myModel;
            }
            this.Cursor = Cursors.Default;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            if (this.radioButton2.Checked)
            {
                dateTimePicker2.Visible = false;

                this.ShangBaoChangDiao.getData(dateTimePicker1.Value.Date);
                this.ShangBaoChangDiao.ModelCalc();
                this.ShangBaoChangDiao.display(webBrowser1.Document);
                this.calModel = this.ShangBaoChangDiao.myModel;
            }
            this.Cursor = Cursors.Default;
        }

        /// <summary>
        ///  其它数据
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            if (this.radioButton3.Checked)
            {
                button2.Visible = true;
                dateTimePicker2.Visible = false;
               this.QiTaShuJu.getData(dateTimePicker1.Value.Date);
               this.QiTaShuJu.ModelCalc();
               this.QiTaShuJu.display(webBrowser1.Document);
            }
            this.Cursor = Cursors.Default;
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            if (this.radioButton4.Checked)
            {
                dateTimePicker2.Visible = true;
                webBrowser1.Document.OpenNew(true);
                tieliangkaohe();
            }
            tabControl1.Visible = radioButton4.Checked;
            PianCha.Visible = radioButton4.Checked;
            this.Cursor = Cursors.Default;
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            if (this.radioButton5.Checked)
            {
                dateTimePicker2.Visible = true;
                button2.Text = "查询条件";
                优质优价炼钢铁 a = new 优质优价炼钢铁();
                a.getData(dateTimePicker1.Value.Date,dateTimePicker2.Value.Date,tiaojian);
                a.display(webBrowser1.Document);
            }
            button2.Visible = radioButton5.Checked;
            this.Cursor = Cursors.Default;
        }

        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            if (this.radioButton6.Checked)
            {
                dateTimePicker2.Visible = true;
                机烧值比较 a = new 机烧值比较();
                a.getData(dateTimePicker1.Value.Date,dateTimePicker2.Value.Date);
                a.display(webBrowser1.Document);
            }
            this.Cursor = Cursors.Default;
        }
    
        private void radioButton7_CheckedChanged(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            if (this.radioButton7.Checked)
            {
                dateTimePicker2.Visible = true;
                优质铁水率 a = new 优质铁水率();
                a.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date);
                a.display(webBrowser1.Document);
            }
            this.Cursor = Cursors.Default;
        }

        private void radioButton8_CheckedChanged(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            if (this.radioButton8.Checked)
            {
                dateTimePicker2.Visible = true;
                生产指标 a = new 生产指标();
                a.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date);
                a.display(webBrowser1.Document);
            }
            this.Cursor = Cursors.Default;
        }

        private void radioButton9_CheckedChanged(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            if (this.radioButton9.Checked)
            {
                dateTimePicker2.Visible = true;
                晚点时间 a = new 晚点时间();
                a.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date);
                a.display(webBrowser1.Document);
            }
            this.Cursor = Cursors.Default;
        }

        private void radioButton10_CheckedChanged(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            if (this.radioButton10.Checked)
            {
                dateTimePicker2.Visible = true;
                大宗原料 a = new 大宗原料();
                a.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date);
                a.display(webBrowser1.Document);
            }
            this.Cursor = Cursors.Default;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            webBrowser1.ShowPrintPreviewDialog();
        }

        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            tieliangkaohe();
            this.Cursor = Cursors.Default;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            if (this.radioButton5.Checked)
            {
                查询条件 a = new 查询条件();
                if (a.ShowDialog() == DialogResult.OK)
                {
                    tiaojian = a.tiaojian;
                    a.Dispose();
                    优质优价炼钢铁 b = new 优质优价炼钢铁();
                    b.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date, tiaojian);
                    b.display(webBrowser1.Document);
                }
            }
            if (this.radioButton12.Checked)
            {
                基准风温 a = new 基准风温();
                if (a.ShowDialog() == DialogResult.OK)
                {
                    风温考核 b = new 风温考核();
                    b.jzfw1 = a.基准风温1;
                    b.jzfw2 = a.基准风温2;
                    b.jzfw3 = a.基准风温3;
                    b.jzfw4 = a.基准风温4;
                    b.jzfw5 = a.基准风温5;
                    a.Dispose();
                    b.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date);
                    b.display(webBrowser1.Document);
                }
            }
            if (this.radioButton3.Checked)
            {
                this.QiTaShuJu.getData(dateTimePicker1.Value.Date);
                this.QiTaShuJu.display(webBrowser1.Document);
            }
            this.Cursor = Cursors.Default;
        }

        private void PianCha_Validated(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            tieliangkaohe();
            this.Cursor = Cursors.Default;
        }


        private void tieliangkaohe()
        {
            ddluciTableAdapter1.FillByRq(this.量化铁考核1.DDLUCI, PianCha.Value, dateTimePicker1.Value.Date, dateTimePicker2.Value.Date, tabControl1.SelectedIndex + 1);
            label1.Text = label4.Text = label6.Text = label8.Text = label10.Text = "共计" + this.量化铁考核1.DDLUCI.Count + "条";

            OracleConnection cn = new OracleConnection();
            cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
            cn.Open();

            OracleCommand cmd = new OracleCommand();
            cmd.Connection = cn;
            cmd.CommandText = "SELECT LUCI FROM DDLUCI WHERE TRUNC(ZDSJ)>=:RQ1 and TRUNC(ZDSJ)<=:RQ2 AND DKSJ IS NOT NULL AND FELIANG IS NULL  AND GAOLU=:GAOLU";
            cmd.Parameters.Add(":RQ1", OracleType.DateTime).Value = dateTimePicker1.Value.Date;
            cmd.Parameters.Add(":RQ2", OracleType.DateTime).Value = dateTimePicker2.Value.Date;
            cmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = tabControl1.SelectedIndex + 1;
            OracleDataReader dr = cmd.ExecuteReader();

            ListBox lb = null;
            switch (tabControl1.SelectedIndex)
            {
                case 0:
                    lb = listBox1;
                    break;
                case 1:
                    lb = listBox2;
                    break;
                case 2:
                    lb = listBox3;
                    break;
                case 3:
                    lb = listBox4;
                    break;
                case 4:
                    lb = listBox5;
                    break;
                case 5:
                    lb = listBox6;
                    break;
            }
            lb.Items.Clear();

            while (dr.Read())
            {
                lb.Items.Add((dr.IsDBNull(0) ? "" : dr.GetString(0)));
            }
            label2.Text = label3.Text = label5.Text = label7.Text = label9.Text = "铁量还未输入的" + lb.Items.Count + "条";
            dr.Close();

            double  a = 0, b = 0, c = 0, d = 0;
            double a1 = 0, b1 = 0, c1 = 0, d1 = 0;
            cmd.CommandText = "SELECT GETBANCI1(ZDSJ) as 班别,COUNT(*) FROM DDLUCI WHERE TRUNC(ZDSJ)>=:RQ1 and TRUNC(ZDSJ)<=:RQ2 AND DKSJ IS NOT NULL  AND GAOLU=:GAOLU group by GETBANCI1(ZDSJ)";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                string banbie = dr.IsDBNull(0) ? "" : dr.GetString(0);
                switch (banbie)
                {
                    case "甲班":
                        a1 = dr.GetDouble(1);
                        break;
                    case "乙班":
                        b1 = dr.GetDouble(1);
                        break;
                    case "丙班":
                        c1 = dr.GetDouble(1);
                        break;
                    case "丁班":
                        d1 = dr.GetDouble(1);
                        break;

                }
            }
            dr.Close();

            cmd.CommandText = "SELECT GETBANCI1(ZDSJ) as 班别,count(*) FROM   DDLUCI  WHERE  (DDLUCI.LFELIANG * (1 - :PC / 100) > (FELIANG+decode(quchu,'铸铁',1,'炼钢',0)*5)) AND (TRUNC(DDLUCI.ZDSJ) >= :SJXIAO) AND (TRUNC(DDLUCI.ZDSJ) <= :SJDA) AND (DDLUCI.GAOLU = :GAOLU) group by GETBANCI1(ZDSJ)";
            cmd.Parameters.Clear();
            cmd.Parameters.Add(":PC", OracleType.Number).Value = PianCha.Value;
            cmd.Parameters.Add(":SJXIAO", OracleType.DateTime).Value = dateTimePicker1.Value.Date;
            cmd.Parameters.Add(":SJDA", OracleType.DateTime).Value = dateTimePicker2.Value.Date;
            cmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = tabControl1.SelectedIndex + 1;
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                string banbie = dr.IsDBNull(0) ? "" : dr.GetString(0);
                switch (banbie)
                {
                    case "甲班":
                        a = dr.GetDouble(1);
                        break;
                    case "乙班":
                        b = dr.GetDouble(1);
                        break;
                    case "丙班":
                        c = dr.GetDouble(1);
                        break;
                    case "丁班":
                        d = dr.GetDouble(1);
                        break;

                }
            }
            dr.Close();

            switch (tabControl1.SelectedIndex)
            {
                case 0:
                    ctl甲1.Text = "甲班： " + a + "/" + a1 + "  " + (a1 == 0 ? "" : ((double)(a * 100 / a1)).ToString("##0.00")) + "%";
                    ctl乙1.Text = "乙班： " + b + "/" + b1 + "  " + (b1 == 0 ? "" : ((double)(b * 100 / b1)).ToString("##0.00")) + "%";
                    ctl丙1.Text = "丙班： " + c + "/" + c1 + "  " + (c1 == 0 ? "" : ((double)(c * 100 / c1)).ToString("##0.00")) + "%";
                    ctl丁1.Text = "丁班： " + d + "/" + d1 + "  " + (d1 == 0 ? "" : ((double)(d * 100 / d1)).ToString("##0.00")) + "%";
                    break;
                case 1:
                    ctl甲2.Text = "甲班： " + a + "/" + a1 + "  " + (a1 == 0 ? "" : ((double)(a * 100 / a1)).ToString("##0.00")) + "%";
                    ctl乙2.Text = "乙班： " + b + "/" + b1 + "  " + (b1 == 0 ? "" : ((double)(b * 100 / b1)).ToString("##0.00")) + "%";
                    ctl丙2.Text = "丙班： " + c + "/" + c1 + "  " + (c1 == 0 ? "" : ((double)(c * 100 / c1)).ToString("##0.00")) + "%";
                    ctl丁2.Text = "丁班： " + d + "/" + d1 + "  " + (d1 == 0 ? "" : ((double)(d * 100 / d1)).ToString("##0.00")) + "%";
                    break;
                case 2:
                    ctl甲3.Text = "甲班： " + a + "/" + a1 + "  " + (a1 == 0 ? "" : ((double)(a * 100 / a1)).ToString("##0.00")) + "%";
                    ctl乙3.Text = "乙班： " + b + "/" + b1 + "  " + (b1 == 0 ? "" : ((double)(b * 100 / b1)).ToString("##0.00")) + "%";
                    ctl丙3.Text = "丙班： " + c + "/" + c1 + "  " + (c1 == 0 ? "" : ((double)(c * 100 / c1)).ToString("##0.00")) + "%";
                    ctl丁3.Text = "丁班： " + d + "/" + d1 + "  " + (d1 == 0 ? "" : ((double)(d * 100 / d1)).ToString("##0.00")) + "%";
                    break;
                case 3:
                    ctl甲4.Text = "甲班： " + a + "/" + a1 + "  " + (a1 == 0 ? "" : ((double)(a * 100 / a1)).ToString("##0.00")) + "%";
                    ctl乙4.Text = "乙班： " + b + "/" + b1 + "  " + (b1 == 0 ? "" : ((double)(b * 100 / b1)).ToString("##0.00")) + "%";
                    ctl丙4.Text = "丙班： " + c + "/" + c1 + "  " + (c1 == 0 ? "" : ((double)(c * 100 / c1)).ToString("##0.00")) + "%";
                    ctl丁4.Text = "丁班： " + d + "/" + d1 + "  " + (d1 == 0 ? "" : ((double)(d * 100 / d1)).ToString("##0.00")) + "%";
                    break;
                case 4:
                    ctl甲5.Text = "甲班： " + a + "/" + a1 + "  " + (a1 == 0 ? "" : ((double)(a * 100 / a1)).ToString("##0.00")) + "%";
                    ctl乙5.Text = "乙班： " + b + "/" + b1 + "  " + (b1 == 0 ? "" : ((double)(b * 100 / b1)).ToString("##0.00")) + "%";
                    ctl丙5.Text = "丙班： " + c + "/" + c1 + "  " + (c1 == 0 ? "" : ((double)(c * 100 / c1)).ToString("##0.00")) + "%";
                    ctl丁5.Text = "丁班： " + d + "/" + d1 + "  " + (d1 == 0 ? "" : ((double)(d * 100 / d1)).ToString("##0.00")) + "%";
                    break;
                case 5:
                    ctl甲6.Text = "甲班： " + a + "/" + a1 + "  " + (a1 == 0 ? "" : ((double)(a * 100 / a1)).ToString("##0.00")) + "%";
                    ctl乙6.Text = "乙班： " + b + "/" + b1 + "  " + (b1 == 0 ? "" : ((double)(b * 100 / b1)).ToString("##0.00")) + "%";
                    ctl丙6.Text = "丙班： " + c + "/" + c1 + "  " + (c1 == 0 ? "" : ((double)(c * 100 / c1)).ToString("##0.00")) + "%";
                    ctl丁6.Text = "丁班： " + d + "/" + d1 + "  " + (d1 == 0 ? "" : ((double)(d * 100 / d1)).ToString("##0.00")) + "%";
                    break;
            }

            cn.Close();
        }

        private void radioButton11_CheckedChanged(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            if (this.radioButton11.Checked)
            {
                dateTimePicker2.Visible = true;
                铁口深度 a = new 铁口深度();
                a.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date);
                a.display(webBrowser1.Document);
            }
            this.Cursor = Cursors.Default;
        }

        private void radioButton12_CheckedChanged(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            if (this.radioButton12.Checked)
            {
                dateTimePicker2.Visible = true;
                button2.Text = "基准风温";
                风温考核 a = new 风温考核();
                a.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date);
                a.display(webBrowser1.Document);
            }
            button2.Visible = radioButton12.Checked;
            this.Cursor = Cursors.Default;
        }

        private void 数据查询窗口_Load(object sender, EventArgs e)
        {
            webBrowser1.DocumentText = "<HTML><body bgcolor=#EEEEFF topmargin=25 leftmargin=5>&nbsp;</BODY></HTML>";
        }

        private void viewCalModel_Click(object sender, EventArgs e)
        {
            if (this.calModel != null)
            {
                CalModelForm form = new CalModelForm(this.calModel);
                form.ShowDialog();
            }
        }

        private void radioButton13_CheckedChanged(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            if (this.radioButton13.Checked)
            {
                dateTimePicker2.Visible = true;
                铁水质量查询条件 a = new 铁水质量查询条件();
                if (a.ShowDialog() == DialogResult.OK)
                {
                    tiaojian = a.tiaojian;
                    a.Dispose();
                    铁水质量查询 b = new 铁水质量查询();
                    b.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date, tiaojian);
                    b.display(webBrowser1.Document);
                }
            
            }
            button3.Visible = radioButton13.Checked;
            this.Cursor = Cursors.Default;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            if (this.radioButton13.Checked)
            {
                dateTimePicker2.Visible = true;
                铁水质量查询条件 a = new 铁水质量查询条件();
                if (a.ShowDialog() == DialogResult.OK)
                {
                    tiaojian = a.tiaojian;
                    a.Dispose();
                    铁水质量查询 b = new 铁水质量查询();
                    b.getData(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date, tiaojian);
                    b.display(webBrowser1.Document);
                }
              
            }
            this.Cursor = Cursors.Default;

        }

     
    }
}