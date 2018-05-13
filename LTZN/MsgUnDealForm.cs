using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace LTZN
{
    public partial class MsgUnDealForm : Form
    {
        public MsgUnDealForm()
        {
            InitializeComponent();
        }

        private void MsgUnDealForm_Load(object sender, EventArgs e)
        {
            loadUnDealMsg();
        }
        /// <summary>
        /// 加载未处理消息
        /// </summary>
        private void loadUnDealMsg()
        {
            listView1.Items.Clear();
            foreach (MsgContent msg in UDPListener.instance.UnDealMsg)
            {
                ListViewItem item = listView1.Items.Add(msg.MsgId, msg.SendUser, 0);
                item.SubItems.Add(msg.SendTime.ToString("yyyy-MM-dd HH:mm"));
                item.SubItems.Add(msg.MsgTyp);
               
            }
        }

        private void listView1_DoubleClick(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                MsgContent msgcontent = null;
                foreach (MsgContent msg in UDPListener.instance.UnDealMsg)
                {
                    if (listView1.SelectedItems.ContainsKey(msg.MsgId))
                    {
                        msgcontent = msg;
                        break;
                    }
                 
                }
                if (msgcontent != null)
                {
                    MsgContentForm frm = new MsgContentForm(msgcontent);
                    frm.ShowDialog();
                }
            }
            loadUnDealMsg();
        }
    }
}