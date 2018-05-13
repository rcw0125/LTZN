using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace LTZN
{
    public partial class 用户管理 : Form
    {
        public 用户管理()
        {
            InitializeComponent();
        }

        private bool vilidItemCheck = true;

        private void 用户管理_Load(object sender, EventArgs e)
        {
            lst_Roles.BeginUpdate();
            lst_AllUsers.BeginUpdate();
            foreach (string role in LtznUserManager.instance.getAllRoles())
            {
                lst_Roles.Items.Add(role);
            }
            lst_AllUsers.DataSource = LtznUserManager.instance.getAllUsers();
            lst_Roles.EndUpdate();
            lst_AllUsers.EndUpdate();
        }

        private void btn_AddUser_Click(object sender, EventArgs e)
        {
                   
            if (txt_ChgPsw1.Text != txt_ChgPsw2.Text)
            {
                MessageBox.Show("输入的两个密码不一致！");
                return;
            }

            if (LtznUserManager.instance.CreateUser(this.txt_UserName.Text, txt_AddPsw1.Text))
                MessageBox.Show("添加完成！");
            else
                MessageBox.Show("添加失败！");
            lst_AllUsers.DataSource = LtznUserManager.instance.getAllUsers();
        }

        private void lst_AllUsers_SelectedIndexChanged(object sender, EventArgs e)
        {
            vilidItemCheck = false;
            lst_Roles.BeginUpdate();
            lbl_Name.Text = lst_AllUsers.Text;

            for (int i = 0; i < lst_Roles.Items.Count; i++)
            {
                if (LtznUserManager.instance.checkUserRole(lst_AllUsers.Text, lst_Roles.Items[i].ToString()))
                    lst_Roles.SetItemChecked(i, true);
                   
                else
                    lst_Roles.SetItemChecked(i, false);
            }
            lst_Roles.EndUpdate();
            vilidItemCheck = true;
        }

        private void btn_ChangePsw_Click(object sender, EventArgs e)
        {
            if (txt_ChgPsw1.Text != txt_ChgPsw2.Text)
            {
                MessageBox.Show("输入的两个密码不一致！");
                return;
            }

            if (LtznUserManager.instance.ChgPsw(lst_AllUsers.Text, txt_ChgPsw1.Text))
                MessageBox.Show("修改完成！");
            else
                MessageBox.Show("修改失败！");
        }

        private void lst_Roles_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (vilidItemCheck)
            {
                if (e.CurrentValue == CheckState.Checked)
                {
                    LtznUserManager.instance.removeRole(lst_AllUsers.Text, lst_Roles.Items[e.Index].ToString());
                }
                else
                {
                    LtznUserManager.instance.addRole(lst_AllUsers.Text, lst_Roles.Items[e.Index].ToString());
                }
            }
        }

        private void btn_DelUser_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("确实要删除此用户吗？", "询问", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
            if (dr == DialogResult.OK)
            {
                LtznUserManager.instance.DelUser(lst_AllUsers.Text);
                lst_AllUsers.DataSource = LtznUserManager.instance.getAllUsers();
            }
        }
    }
}