using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace LTZN.数据分析
{
    public partial class 原料分类 : Form
    {
        public 原料分类()
        {
            InitializeComponent();
        }

        private void 原料分类_Load(object sender, EventArgs e)
        {
            // TODO: 这行代码将数据加载到表“dataSet2.YLFL”中。您可以根据需要移动或移除它。
            this.yLFLTableAdapter.Fill(this.dataSet2.YLFL);

        }

      
        private void button1_Click(object sender, EventArgs e)
        {
            this.yLFLTableAdapter.Update(this.dataSet2.YLFL);
            MessageBox.Show("保存完成");
        }
    }
}
