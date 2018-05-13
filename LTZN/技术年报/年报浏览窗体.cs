using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace LTZN.技术年报
{
    public partial class 年报浏览窗体 : Form
    {
        public 年报浏览窗体()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            webBrowser1.Navigate("http://" + Properties.Settings.Default.服务器 + "/LTWeb/nianbao.aspx?nian=" + this.comboBox1.Text);
            
        }
    }
}