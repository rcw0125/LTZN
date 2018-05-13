using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace LTZN.质量日报
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: 这行代码将数据加载到表“dataSet3.DDYL”中。您可以根据需要移动或移除它。
            this.dDYLTableAdapter.Fill(this.dataSet3.DDYL);

        }
    }
}
