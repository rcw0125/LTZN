using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace LTZN
{
    public partial class NameDialog : Form
    {
        public NameDialog()
        {
            InitializeComponent();
        }

        private void NameDialog_Load(object sender, EventArgs e)
        {
           // txtName.Focus();

        }

        public string Result
        {
            get
            {
                return txtName.Text;
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
