using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using C1.Win.C1Chart;

namespace LTZN.数据分析
{
    public partial class 出铁数据 : Form
    {
        public 出铁数据()
        {
            InitializeComponent();
        }

        private void 产量_Load(object sender, EventArgs e)
        {
            this.sjBegin.Value = DateTime.Today.AddDays(-2);
            this.sjEnd.Value = DateTime.Now;
        }
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            if (r1.RadioChecked && (sjEnd.Value < sjBegin.Value))
            {
                MessageBox.Show("结束时间应该比开始时间大！");
                return;
            }
         
            switch (this.toolStrip_gaolu.Text)
            {
                case "1高炉":
                    if (r1.RadioChecked)
                        this.dDLUCITableAdapter.FillByGLSJ(this.chuTieDataSet.DDLUCI, 1, this.sjBegin.Value, this.sjEnd.Value);
                    else
                        this.dDLUCITableAdapter.FillByGLLC(this.chuTieDataSet.DDLUCI, 1, this.toolStrip_Luci1.Text, this.toolStrip_Luci2.Text);
                    break;
      
                case "2高炉":
                    if (r1.RadioChecked)
                        this.dDLUCITableAdapter.FillByGLSJ(this.chuTieDataSet.DDLUCI, 2, this.sjBegin.Value, this.sjEnd.Value);
                    else
                        this.dDLUCITableAdapter.FillByGLLC(this.chuTieDataSet.DDLUCI, 2, this.toolStrip_Luci1.Text, this.toolStrip_Luci2.Text);
                    break;

                case "3高炉":
                    if (r1.RadioChecked)
                        this.dDLUCITableAdapter.FillByGLSJ(this.chuTieDataSet.DDLUCI, 3, this.sjBegin.Value, this.sjEnd.Value);
                    else
                        this.dDLUCITableAdapter.FillByGLLC(this.chuTieDataSet.DDLUCI, 3, this.toolStrip_Luci1.Text, this.toolStrip_Luci2.Text);
                    break;

                case "4高炉":
                    if (r1.RadioChecked)
                        this.dDLUCITableAdapter.FillByGLSJ(this.chuTieDataSet.DDLUCI, 4, this.sjBegin.Value, this.sjEnd.Value);
                    else
                        this.dDLUCITableAdapter.FillByGLLC(this.chuTieDataSet.DDLUCI, 4, this.toolStrip_Luci1.Text, this.toolStrip_Luci2.Text);
                    break;

                case "5高炉":
                    if (r1.RadioChecked)
                        this.dDLUCITableAdapter.FillByGLSJ(this.chuTieDataSet.DDLUCI, 5, this.sjBegin.Value, this.sjEnd.Value);
                    else
                        this.dDLUCITableAdapter.FillByGLLC(this.chuTieDataSet.DDLUCI, 5, this.toolStrip_Luci1.Text, this.toolStrip_Luci2.Text);
                    break;

                case "6高炉":
                    if (r1.RadioChecked)
                        this.dDLUCITableAdapter.FillByGLSJ(this.chuTieDataSet.DDLUCI, 6, this.sjBegin.Value, this.sjEnd.Value);
                    else
                        this.dDLUCITableAdapter.FillByGLLC(this.chuTieDataSet.DDLUCI,6, this.toolStrip_Luci1.Text, this.toolStrip_Luci2.Text);
                    break;
                default :
                    if (r1.RadioChecked)
                        this.dDLUCITableAdapter.FillBySJ(this.chuTieDataSet.DDLUCI, this.sjBegin.Value, this.sjEnd.Value);
                    else
                        this.dDLUCITableAdapter.FillByLC(this.chuTieDataSet.DDLUCI, this.toolStrip_Luci1.Text, this.toolStrip_Luci2.Text);
                    break;
            }

            this.Cursor = Cursors.Default;

        }

        private void r1_CheckedChanged(object sender, EventArgs e)
        {
            if (r1.RadioChecked)
            {
                sjBegin.Visible = true;
                sjEnd.Visible = true;
                toolStrip_Luci1.Visible = false;
                toolStrip_Luci2.Visible = false;
                toolStripLabel1.Text = "开始时间";
                toolStripLabel2.Text = "结束时间";
            }
            else
            {
                sjBegin.Visible = false;
                sjEnd.Visible = false;
                toolStrip_Luci1.Visible = true;
                toolStrip_Luci2.Visible = true;
                toolStripLabel1.Text = "开始炉次";
                toolStripLabel2.Text = "结束炉次";
            }
        }

        private void toolStrip_Luci1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                toolStrip_Luci2.Focus();
        }

        private void toolStrip_Luci2_KeyDown(object sender, KeyEventArgs e)
        {
        }

        private void c1Chart1_Load(object sender, EventArgs e)
        {

        }

        private void 出铁数据_Load(object sender, EventArgs e)
        {
            this.sjBegin.Value = DateTime.Today.AddDays(-2);
            this.sjEnd.Value = DateTime.Now;
        }
    }
}