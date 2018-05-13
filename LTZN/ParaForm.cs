using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace LTZN
{
    public partial class ParaForm : Form
    {
        public ParaForm()
        {
            InitializeComponent();
        }


        private void btnOk_Click(object sender, EventArgs e)
        {

            LTZN.Data.SysParameter.SetParameter("单位负责人",txtLeader.Text);
            LTZN.Data.SysParameter.SetParameter("日报负责人", txtActor1.Text);
            LTZN.Data.SysParameter.SetParameter("月报负责人", txtActor2.Text);
            this.Close();
        }

        private void ParaForm_Load(object sender, EventArgs e)
        {
            String leader = LTZN.Data.SysParameter.GetStrParameter("单位负责人", "陈学清");
            String dayReportAuthor = LTZN.Data.SysParameter.GetStrParameter("日报负责人", "刘红芬");
            String monthReportAuthor = LTZN.Data.SysParameter.GetStrParameter("月报负责人", "王文英");
            txtLeader.Text = leader;
            txtActor1.Text = dayReportAuthor;
            txtActor2.Text = monthReportAuthor;

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
