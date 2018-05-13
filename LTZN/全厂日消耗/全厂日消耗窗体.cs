using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;
using C1.Win.C1Input;

namespace LTZN.全厂日消耗
{
    public partial class 全厂日消耗窗体 : Form
    {
        private bool canUpdate = false;

        public 全厂日消耗窗体()
        {
            InitializeComponent();
            LtznUserManager.instance.RegisterHandler(this, 全厂日消耗窗体_UserChange);
        }
        private void getData(DateTime riqi)
        { 
            if (canUpdate)
            {
                this.flex原料.FinishEditing();
                this.bindingSource1.CurrencyManager.EndCurrentEdit();
                this.bindingSource2.CurrencyManager.EndCurrentEdit();
                this.全厂日消耗TableAdapter.Update(this.全厂日消耗数据集1.全厂日消耗);
                this.日原料成分TableAdapter.Update(this.全厂日消耗数据集1.日原料成分);
            }
            this.全厂日消耗TableAdapter.FillByRq(this.全厂日消耗数据集1.全厂日消耗, riqi);
            this.日原料成分TableAdapter.FillByRq(this.全厂日消耗数据集1.日原料成分, riqi);
            if (this.全厂日消耗数据集1.全厂日消耗.Count == 0)
            {
                this.全厂日消耗数据集1.全厂日消耗.Add全厂日消耗Row(riqi);
            }
            if (this.全厂日消耗数据集1.日原料成分.Count != 11)
            {
                List<string> yuanliao = new List<string>();
                yuanliao.Add("机烧");
                yuanliao.Add("竖炉球");
                yuanliao.Add("PB块");
                yuanliao.Add("FMG块");
                yuanliao.Add("锰矿");
                yuanliao.Add("钛球");
                yuanliao.Add("瓦斯灰");
                yuanliao.Add("其它熟料");
                yuanliao.Add("其它生料");
                yuanliao.Add("本溪矿");
                yuanliao.Add("印度球");

                foreach (var item in yuanliao)
                {
                     全厂日消耗数据集.日原料成分Row rowdata= this.全厂日消耗数据集1.日原料成分.FindByP01日期P02名称(riqi, item);
                 if (rowdata == null)
                 {
                     this.全厂日消耗数据集1.日原料成分.Add日原料成分Row(riqi, item);
                 }  
                }
              
                //this.全厂日消耗数据集1.日原料成分.Add日原料成分Row(riqi, "机烧");
                //this.全厂日消耗数据集1.日原料成分.Add日原料成分Row(riqi, "竖炉球");
                //this.全厂日消耗数据集1.日原料成分.Add日原料成分Row(riqi, "其它熟料");
                //this.全厂日消耗数据集1.日原料成分.Add日原料成分Row(riqi, "其它生料");
                //this.全厂日消耗数据集1.日原料成分.Add日原料成分Row(riqi, "本溪矿");
                //this.全厂日消耗数据集1.日原料成分.Add日原料成分Row(riqi, "印度球");
                //this.全厂日消耗数据集1.日原料成分.Add日原料成分Row(riqi, "PB块");
                //this.全厂日消耗数据集1.日原料成分.Add日原料成分Row(riqi, "FMG块");
                //this.全厂日消耗数据集1.日原料成分.Add日原料成分Row(riqi, "锰矿");
                //this.全厂日消耗数据集1.日原料成分.Add日原料成分Row(riqi, "钛球"); 
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            getData(this.dateTimePicker1.Value.Date);
        }

        private void 全厂日消耗窗体_Load(object sender, EventArgs e)
        {
            this.dateTimePicker1.Value = DateTime.Today;
        }

        private void 全厂日消耗窗体_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (canUpdate)
            {
                this.flex原料.FinishEditing();
                this.bindingSource1.CurrencyManager.EndCurrentEdit();
                this.bindingSource2.CurrencyManager.EndCurrentEdit();
                this.全厂日消耗TableAdapter.Update(this.全厂日消耗数据集1.全厂日消耗);
                this.日原料成分TableAdapter.Update(this.全厂日消耗数据集1.日原料成分);
            }
        }

        private void c1NumericEdit27_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Delete:
                    ((C1NumericEdit)sender).Value = System.DBNull.Value;
                    break;
                case Keys.Enter:
                    ((Control)sender).Parent.Parent.SelectNextControl((Control)sender, true, true, true, false);
                    break;
                case Keys.Down:
                    ((Control)sender).Parent.Parent.SelectNextControl((Control)sender, true, true, true, false);
                    break;
                case Keys.Up:
                    ((Control)sender).Parent.Parent.SelectNextControl((Control)sender, false, true, true, false);
                    break;
            }
        }

        private void c1TextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Enter:
                    ((Control)sender).Parent.Parent.SelectNextControl((Control)sender, true, true, true, false);
                    break;
                case Keys.Down:
                    ((Control)sender).Parent.Parent.SelectNextControl((Control)sender, true, true, true, false);
                    break;
                case Keys.Up:
                    ((Control)sender).Parent.Parent.SelectNextControl((Control)sender, false, true, true, false);
                    break;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OracleConnection cn = new OracleConnection();
            cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
            cn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = cn;
            cmd.CommandText = "select avg(shuif) from ddjt where trunc(sj)=:RQ and mc in ('自产焦夜','自产焦白','自产焦中')";
            cmd.Parameters.Add(":RQ", OracleType.DateTime).Value = dateTimePicker1.Value.Date;
            try { c1NumericEdit1.Value = cmd.ExecuteScalar(); }
            catch { c1NumericEdit1.Value = System.DBNull.Value; }
            cmd.CommandText = "select avg(shuif) from ddjt where trunc(sj)=:RQ  and mc='外进焦'";
            try { c1NumericEdit2.Value = cmd.ExecuteScalar(); }
            catch { c1NumericEdit2.Value = System.DBNull.Value; }
            cn.Close();
        }

        private void 全厂日消耗窗体_Leave(object sender, EventArgs e)
        {
            if (canUpdate)
            {
                this.flex原料.FinishEditing();
                this.bindingSource1.CurrencyManager.EndCurrentEdit();
                this.bindingSource2.CurrencyManager.EndCurrentEdit();
                this.全厂日消耗TableAdapter.Update(this.全厂日消耗数据集1.全厂日消耗);
                this.日原料成分TableAdapter.Update(this.全厂日消耗数据集1.日原料成分);
            }
        }

        void 全厂日消耗窗体_UserChange(LtznUser ltznUser)
        {
            if (ltznUser != null && ltznUser.IsInRole("统计"))
            {
                canUpdate = true;
                flex原料.AllowEditing = true;
                button1.Enabled = true;
            }
            else
            {
                canUpdate = false;
                flex原料.AllowEditing = false;
                button1.Enabled = false;
            }
        }
    }
}