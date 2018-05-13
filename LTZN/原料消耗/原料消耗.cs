using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace LTZN.原料消耗
{
    public partial class 原料消耗 : Form
    {
        public 原料消耗()
        {
            InitializeComponent();
        }

        private void 原料消耗保存()
        {
            this.Validate();
            this.原料消耗BindingSource.EndEdit();
            this.原料消耗TableAdapter.Update(this.原料消耗数据集.原料消耗);

        }

        private void 原料消耗_Load(object sender, EventArgs e)
        {
            foreach (string s1 in Properties.Settings.Default.熟料分类)
            {
                shuliao.Items.Add(s1);
            }
            foreach (string s2 in Properties.Settings.Default.生料分类)
            {
                shengliao.Items.Add(s2);
            }
            getData(DateTime.Today, 1);

        }

        private void getData(DateTime riqi,decimal gaolu)
        {
            this.原料消耗TableAdapter.FillByPk(this.原料消耗数据集.原料消耗, riqi, gaolu);
            if (this.原料消耗数据集.原料消耗.Count == 0)
            {
                this.原料消耗数据集.原料消耗.Add原料消耗Row(riqi, gaolu);
            }
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            this.原料消耗保存();
            this.getData(this.dateTimePicker1.Value.Date, Convert.ToDecimal(this.comboBox1.Text));
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.原料消耗保存();
            this.getData(this.dateTimePicker1.Value.Date, Convert.ToDecimal(this.comboBox1.Text));
        }

        private void 原料消耗_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.原料消耗保存();
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }
}