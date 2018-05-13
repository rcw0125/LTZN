using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace LTZN
{
    public partial class 出铁方案 : Form
    {
        public 出铁方案()
        {
            InitializeComponent();
        }

        private void 出铁方案_Load(object sender, EventArgs e)
        {
            // TODO: 这行代码将数据加载到表“ckfa.CHUTIEFANGAN”中。您可以根据需要移动或移除它。
            this.cHUTIEFANGANTableAdapter.Fill(this.ckfa.CHUTIEFANGAN);

        }
   
        private void 保存_Click(object sender, EventArgs e)
        {
           
                this.cHUTIEFANGANTableAdapter.Update(this.ckfa.CHUTIEFANGAN);
                MessageBox.Show("保存完成");
            
           
        }

        private void dataGridView1_DataError_1(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("不能输入空值");
        }
   
    }
}
