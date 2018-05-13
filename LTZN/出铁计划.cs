using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace LTZN
{
    public partial class 出铁计划 : Form
    {
        public 出铁计划()
        {
            InitializeComponent();
        }

        private void 出铁计划_Load(object sender, EventArgs e)
        {
            // TODO: 这行代码将数据加载到表“ctjh.CHUTIEJIHUA”中。您可以根据需要移动或移除它。
            this.cHUTIEJIHUATableAdapter.Fill(this.ctjh.CHUTIEJIHUA);

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.cHUTIEJIHUATableAdapter.Update(this.ctjh.CHUTIEJIHUA);
            MessageBox.Show("保存完成");
        }

        private void dataGridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("不能输入空值");
        }
    }
}
