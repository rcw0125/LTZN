using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace LTZN.数据分析
{
    public partial class 趋势参数 : Form
    {
        public 趋势参数()
        {
            InitializeComponent();
        }

        private void 趋势参数_Load(object sender, EventArgs e)
        {
            // TODO: 这行代码将数据加载到表“dataSet1.LT_QSPARA”中。您可以根据需要移动或移除它。
            this.lT_QSPARATableAdapter.Fill(this.dataSet1.LT_QSPARA);

        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            this.lT_QSPARATableAdapter.Update(this.dataSet1.LT_QSPARA);
        }
    }
}
