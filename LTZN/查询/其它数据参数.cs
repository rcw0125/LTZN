using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Security.Principal;

namespace LTZN.查询
{
    public partial class 其它数据参数 : Form
    {
        public 其它数据参数()
        {
            InitializeComponent();
            LtznUserManager.instance.RegisterHandler(this, instance_UserChanged);
        }

        public double Precent1
        {
            get { return Convert.ToDouble(txtPrecent1.Text); }
            set { txtPrecent1.Text = value.ToString(); }
        }

        public double Amount1
        {
            get { return Convert.ToDouble(txtAmount1.Text); }
            set { txtAmount1.Text = value.ToString(); }
        }

        public bool isPrecent1
        {
            get 
            {
                return this.rdoPrecent1.Checked;
            }
            set { rdoPrecent1.Checked = value; }
        }

        public bool isAccount1
        {
            get
            {
                return this.rdoAmount1.Checked;
            }
            set { rdoAmount1.Checked = value; }
        }



        public double Precent2
        {
            get { return Convert.ToDouble(txtPrecent2.Text); }
            set { txtPrecent2.Text = value.ToString(); }
        }

        public double Amount2
        {
            get { return Convert.ToDouble(txtAmount2.Text); }
            set { txtAmount2.Text = value.ToString(); }
        }

        public bool isPrecent2
        {
            get
            {
                return this.rdoPrecent2.Checked;
            }
            set { rdoPrecent2.Checked = value; }
        }


        public bool isAccount2
        {
            get
            {
                return this.rdoAmount2.Checked;
            }
            set { rdoAmount2.Checked = value; }
        }

        public double Precent3
        {
            get { return Convert.ToDouble(txtPrecent3.Text); }
            set { txtPrecent3.Text = value.ToString(); }
        }

        public double Amount3
        {
            get { return Convert.ToDouble(txtAmount3.Text); }
            set { txtAmount3.Text = value.ToString(); }
        }

        public bool isPrecent3
        {
            get
            {
                return this.rdoPrecent3.Checked;
            }
            set { rdoPrecent3.Checked = value; }

        }

        public bool isAccount3
        {
            get
            {
                return this.rdoAmount3.Checked;
            }
            set { rdoAmount3.Checked = value; }
        }

        public double Precent4
        {
            get { return Convert.ToDouble(txtPrecent4.Text); }
            set { txtPrecent4.Text = value.ToString(); }
        }

        public double Amount4
        {
            get { return Convert.ToDouble(txtAmount4.Text); }
            set { txtAmount4.Text = value.ToString(); }
        }

        public bool isPrecent4
        {
            get
            {
                return this.rdoPrecent4.Checked;
            }
            set { rdoPrecent4.Checked = value; }
        }

        public bool isAccount4
        {
            get
            {
                return this.rdoAmount4.Checked;
            }
            set { rdoAmount4.Checked = value; }
        }

        public double Precent5
        {
            get { return Convert.ToDouble(txtPrecent5.Text); }
            set { txtPrecent5.Text = value.ToString(); }
        }

        public double Amount5
        {
            get { return Convert.ToDouble(txtAmount5.Text); }
            set { txtAmount5.Text = value.ToString(); }
        }

        public bool isPrecent5
        {
            get
            {
                return this.rdoPrecent5.Checked;
            }
            set { rdoPrecent5.Checked = value; }
        }

        public bool isAccount5
        {
            get
            {
                return this.rdoAmount5.Checked;
            }
            set { rdoAmount5.Checked = value; }
        }

        public bool isSave
        {
            get {
                return this.rdoSave.Checked;
            }
        }


        private void 其它数据参数_Load(object sender, EventArgs e)
        {
          
        }
        void instance_UserChanged(LtznUser ltznUser)
        {
            IPrincipal p = LtznUserManager.instance.CurrentUser;
            rdoLinshi.Enabled = false;
            rdoSave.Enabled = false;
            if (p != null)
            {
                if (p.IsInRole("管理员"))
                {
                    rdoLinshi.Enabled = true;
                    rdoSave.Enabled = true;
                }
                return;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.OK;
            this.Close();

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void 其它数据参数_FormClosed(object sender, FormClosedEventArgs e)
        {
        }
    }
}
