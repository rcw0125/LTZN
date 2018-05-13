using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;

namespace LTZN
{
    public partial class MsgContentForm : Form
    {
        private MsgContent msgContent = null;
        public MsgContentForm(MsgContent msgContent)
        {
            InitializeComponent();
            this.msgContent = msgContent;
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            UDPListener.instance.DealMsg(msgContent);
            this.Close();
        }

        private void MsgContentForm_Load(object sender, EventArgs e)
        {
            lSendUser.Text = msgContent.SendUser;
            lSendTime.Text = msgContent.SendTime.ToString("yyyy-MM-dd HH:mm:ss");
            lMsgTyp.Text = msgContent.MsgTyp;
            txtContent.Text = msgContent.SendContent;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MsgContent msgContentTemp = new MsgContent();
            msgContentTemp.SendUser = LtznUserManager.instance.CurrentUser.Identity.Name;
            msgContentTemp.ToUser = this.msgContent.SendUser;
            msgContentTemp.MsgTyp = "回复";
            msgContentTemp.SendContent = txtHuifu.Text.TrimEnd();
            msgContentTemp.ByMsg = this.msgContent.MsgId;
            UDPListener.instance.sendMsg(msgContentTemp);
            UDPListener.instance.DealMsg(this.msgContent);
            this.Close();
        }
    }
}