using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;
using C1.Win.C1Input;
using LTZN.Data;
using LTZN.Repository;

namespace LTZN.技术月报
{
    public partial class 月报调整 : Form
    {
        private bool 更改 = false;
        private string p年 = "";
        private string p月 = "";
        private bool 更新 = true;

        private MonthEnergy _MonthEnergy = new MonthEnergy();
        public MonthEnergy MonthEnergy
        {
            get { return _MonthEnergy; }
            set {
                if (_MonthEnergy != value)
                {
                    _MonthEnergy = value;
                }
            }
        }

        public 月报调整()
        {
            InitializeComponent();
         
        }

        private void 月报调整_Load(object sender, EventArgs e)
        {   
            更新 = false;
            ctl年.Text = DateTime.Today.Year + "年";
            ctl月.Text = DateTime.Today.Month + "月";
            更新 = true;
            getData();
        }

        private void save()
        {
            try
            {
                if (!LtznUserManager.instance.CurrentUser.IsInRole("统计"))
                    return;

                if (!更改) return;
                OracleConnection cn = new OracleConnection();
                cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
                cn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = cn;
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "UPDATE每月矿";
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    cmd.Parameters.Clear();
                    cmd.Parameters.Add("Inian", OracleType.Int32).Value = Convert.ToInt32(p年.Replace("年", ""));
                    cmd.Parameters.Add("Iyue", OracleType.Int32).Value = Convert.ToInt32(p月.Replace("月", ""));
                    cmd.Parameters.Add("Imc", OracleType.VarChar, 20).Value = dataGridView1[0, i].Value;
                    cmd.Parameters.Add("Izhongliang", OracleType.Number).Value = dataGridView1[1, i].Value == null ? DBNull.Value : (object)dataGridView1[1, i].Value;
                    cmd.Parameters.Add("Itfe", OracleType.Number).Value = dataGridView1[2, i].Value == null ? DBNull.Value : (object)dataGridView1[2, i].Value;
                    cmd.Parameters.Add("Ifeo", OracleType.Number).Value = dataGridView1[3, i].Value == null ? DBNull.Value : (object)dataGridView1[3, i].Value;
                    cmd.Parameters.Add("Isio2", OracleType.Number).Value = dataGridView1[4, i].Value == null ? DBNull.Value : (object)dataGridView1[4, i].Value;
                    cmd.Parameters.Add("Icao", OracleType.Number).Value = dataGridView1[5, i].Value == null ? DBNull.Value : (object)dataGridView1[5, i].Value;
                    cmd.Parameters.Add("Imgo", OracleType.Number).Value = dataGridView1[6, i].Value == null ? DBNull.Value : (object)dataGridView1[6, i].Value;
                    cmd.Parameters.Add("Ir", OracleType.Number).Value = dataGridView1[7, i].Value == null ? DBNull.Value : (object)dataGridView1[7, i].Value;
                    cmd.Parameters.Add("Iliu", OracleType.Number).Value = dataGridView1[8, i].Value == null ? DBNull.Value : (object)dataGridView1[8, i].Value;
                    cmd.ExecuteNonQuery();
                }
                cn.Close();
                MonthEnergyRepository rp = new MonthEnergyRepository();
                rp.Save(MonthEnergy);
                更改 = false;
            }
            catch
            {

                return;
            }
            
        }

        private void getData()
        {
            List<string> yuanliao = new List<string>();
            //yuanliao.Add("机烧");
            //yuanliao.Add("竖炉球");
            //yuanliao.Add("其它熟料");
            //yuanliao.Add("本溪矿");
            //yuanliao.Add("其它生料");
            //yuanliao.Add("瓦斯灰");
            //yuanliao.Add("PB块");
            //yuanliao.Add("FMG块");
            //yuanliao.Add("锰矿");
            //yuanliao.Add("钛球");

            yuanliao.Add("机烧");
            yuanliao.Add("球团矿");
            yuanliao.Add("国内球团矿");
            yuanliao.Add("进口球团矿");
            yuanliao.Add("PB块");
            yuanliao.Add("纽曼块");
            yuanliao.Add("其它块矿");
            yuanliao.Add("高钛球团矿");
            yuanliao.Add("高品位钛球");
            yuanliao.Add("瓦斯灰");

            dataGridView1.Rows.Clear();
            foreach (var item in yuanliao)
            {
                this.dataGridView1.Rows.Add(item);   
            }
     
            OracleConnection cn = new OracleConnection();
            cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
            cn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = cn;
            cmd.Parameters.Add(":nian", OracleType.Int32).Value = Convert.ToInt32(ctl年.Text.Replace("年", ""));
            cmd.Parameters.Add(":yue", OracleType.Int32).Value = Convert.ToInt32(ctl月.Text.Replace("月", ""));
            cmd.CommandText = "select 名称,重量,TFE,FEO,SIO2,CAO,MGO,R,S from 每月矿 where 年=:nian and 月=:yue";
            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(0))
                {
                    string mc = dr.GetString(0);
                    int row= yuanliao.IndexOf(mc);
                    if (row >= 0)
                    {
                        for (int col = 1; col <= 8; col++)
                        {
                             dataGridView1[col, row].Value = dr[col];
                        }
                    }
                    
                }

            }
            dr.Close();
            dataGridView1.Refresh();
            cn.Close();
            MonthEnergyRepository rp = new MonthEnergyRepository();
            int nian = Convert.ToInt32(ctl年.Text.Replace("年", ""));
            int yue = Convert.ToInt32(ctl月.Text.Replace("月", ""));
            this.MonthEnergy = rp.GetMonthEnergy(nian, yue);
            bindingSource1.DataSource = this.MonthEnergy;
            bindingSource1.ResetBindings(false);
            p年 = ctl年.Text;
            p月 = ctl月.Text;
            更改 = false;
        }

        private void ctl月_SelectedValueChanged(object sender, EventArgs e)
        {
            if (更新)
            {
                save();
                getData();
            }
        }

        private void ctl年_SelectedValueChanged(object sender, EventArgs e)
        {
            if (更新)
            {
                save();
                getData();
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            更改 = true;
        }

        private void c1NumericEdit1_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Delete:
                    ((C1NumericEdit)sender).Value = System.DBNull.Value;
                    break;
                case Keys.Enter:
                    ((Control)sender).Parent.SelectNextControl((Control)sender, true, true, true, false);
                    break;
                case Keys.Down:
                    ((Control)sender).Parent.SelectNextControl((Control)sender, true, true, true, false);
                    break;
                case Keys.Up:
                    ((Control)sender).Parent.SelectNextControl((Control)sender, false, true, true, false);
                    break;
            }
        }

        private void c1NumericEdit12_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Delete:
                    ((C1NumericEdit)sender).Value = System.DBNull.Value;
                    break;
                case Keys.Enter:
                    c1NumericEdit13.Focus();
                    break;
                case Keys.Down:
                    c1NumericEdit13.Focus();
                    break;
                //case Keys.Up:
                //    c1NumericEdit11.Focus();
                //    break;
            }
        }

        private void c1NumericEdit19_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Delete:
                    ((C1NumericEdit)sender).Value = System.DBNull.Value;
                    break;
                case Keys.Enter:
                    c1NumericEdit20.Focus();
                    break;
                case Keys.Down:
                    c1NumericEdit20.Focus();
                    break;
                case Keys.Up:
                    c1NumericEdit18.Focus();
                    break;
            }
        }

        private void c1NumericEdit23_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Delete:
                    ((C1NumericEdit)sender).Value = System.DBNull.Value;
                    break;
                //case Keys.Enter:
                //    c1NumericEdit24.Focus();
                //    break;
                //case Keys.Down:
                //    c1NumericEdit24.Focus();
                //    break;
                case Keys.Up:
                    c1NumericEdit22.Focus();
                    break;
            }
        }

        private void 月报调整_Leave(object sender, EventArgs e)
        {
            save();
        }

        private void c1NumericEdit1_TextChanged(object sender, EventArgs e)
        {
            更改 = true;
        }
    }
}