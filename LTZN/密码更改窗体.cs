using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;

namespace LTZN
{
    public partial class 密码更改窗体 : Form
    {
        public string yonghu = "";

        public 密码更改窗体()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Trim() == textBox2.Text.Trim())
            {
                OracleConnection cn = new OracleConnection();
                cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
                cn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = cn;
                cmd.CommandText = "update users set mima='" + textBox1.Text.Trim() + "' where yonghu='" + yonghu.Trim() + "' and mima='" + mimaCtl.Text.Trim() + "'";
                if (cmd.ExecuteNonQuery()==0)
                    MessageBox.Show("修改密码失败，请检查旧密码！");
                else
                    MessageBox.Show("修改密码成功！");
                cn.Close();
            }
            else
                MessageBox.Show("请确认两次输入的密码是否一致！");
        }
    }
}