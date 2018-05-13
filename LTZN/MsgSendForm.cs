using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.Data.OracleClient;
using System.Net.Sockets;

namespace LTZN
{
    public partial class MsgSendForm : Form
    {
        private const int listenPort = 11000;

        public MsgSendForm()
        {
            InitializeComponent();
        }

        private void MsgSendForm_Load(object sender, EventArgs e)
        {
            foreach (string user in LtznUserManager.instance.getAllUsers())
            {
                lstAllUser.Items.Add(user);
            }
        }

        private void btnAddOne_Click(object sender, EventArgs e)
        {
            string user = lstAllUser.Text.Trim();
            if (!String.IsNullOrEmpty(user) && !lstToUsers.Items.Contains(user))
            {
                lstToUsers.Items.Add(user);
            }
        }

        private void btnAddAll_Click(object sender, EventArgs e)
        {
            lstToUsers.Items.Clear();
            foreach (string user in lstAllUser.Items)
            {
                lstToUsers.Items.Add(user);
            }
        }

        private void btnDelAll_Click(object sender, EventArgs e)
        {
            lstToUsers.Items.Clear();
        }

        private void btnDelOne_Click(object sender, EventArgs e)
        {
            lstToUsers.Items.Remove(lstToUsers.SelectedItem);
        }

        private void btnSend_Click(object sender, EventArgs e)
        {


            string tousers = "";
            foreach (string user in lstToUsers.Items)
            {
                tousers += user + ",";
            }
            if (tousers.Length > 0)
            {

                tousers = tousers.Remove(tousers.Length - 1, 1);

                string typ = "";
                foreach (Control control in grpTyps.Controls)
                {
                    if (control.GetType() == typeof(RadioButton))
                    {
                        if (((RadioButton)control).Checked)
                        {
                            typ = ((RadioButton)control).Text;
                            break;
                        }

                    }
                }

                MsgContent msgContent = new MsgContent();
                msgContent.SendUser = LtznUserManager.instance.CurrentUser.Identity.Name;
                msgContent.ToUser = tousers;
                msgContent.MsgTyp = typ;
                msgContent.SendContent = txtContent.Text.TrimEnd();
                UDPListener.instance.sendMsg(msgContent);
                txtContent.Text = "";
            }
            else
            {
                MessageBox.Show("没有目标用户。");
            }

        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}