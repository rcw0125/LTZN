using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace LTZN.查询
{
    public partial class 铁水质量查询条件 : Form
    {
        public string tiaojian = "";
        public 铁水质量查询条件()
        {
            InitializeComponent();
        }

        private void 铁水质量查询条件_Load(object sender, EventArgs e)
        {
            

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (SCheck.Checked == true)
            {
                tiaojian = "FES<=0.02";
                if (PCheck.Checked == true)
                {
                    tiaojian = tiaojian + " and FEP<=0.09";
                    if (TICheck.Checked == true)
                    {
                        tiaojian = tiaojian + " and FETI<=0.055";
                    }
                    else
                    {
                       
                    }
                }
                else
                {
                    if (TICheck.Checked == true)
                    {
                        tiaojian = tiaojian + " and FETI<=0.055";
                    }
                    else
                    {
                      
                    }
                }
            }
            else
            {
                if (PCheck.Checked == true)
                {
                    tiaojian = tiaojian + " FEP<=0.09";
                    if (TICheck.Checked == true)
                    {
                        tiaojian = tiaojian + " and FETI<=0.055";
                    }
                    else
                    {
                      
                    }
                }
                else
                {
                    if (TICheck.Checked == true)
                    {
                        tiaojian = tiaojian + " FETI<=0.055";
                    }
                    else
                    {
                        MessageBox.Show("请选择查询条件");
                    }
                }

            }
            this.DialogResult = DialogResult.OK;
            this.Close();

        }
    }
}
