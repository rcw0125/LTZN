namespace LTZN
{
    partial class 用户管理
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.panel1 = new System.Windows.Forms.Panel();
            this.label8 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txt_AddPsw2 = new System.Windows.Forms.TextBox();
            this.txt_AddPsw1 = new System.Windows.Forms.TextBox();
            this.btn_AddUser = new System.Windows.Forms.Button();
            this.txt_UserName = new System.Windows.Forms.TextBox();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.lst_AllUsers = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.lst_Roles = new System.Windows.Forms.CheckedListBox();
            this.label4 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.btn_DelUser = new System.Windows.Forms.Button();
            this.lbl_Name = new System.Windows.Forms.Label();
            this.btn_ChangePsw = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.txt_ChgPsw2 = new System.Windows.Forms.TextBox();
            this.txt_ChgPsw1 = new System.Windows.Forms.TextBox();
            this.panel1.SuspendLayout();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.label8);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.txt_AddPsw2);
            this.panel1.Controls.Add(this.txt_AddPsw1);
            this.panel1.Controls.Add(this.btn_AddUser);
            this.panel1.Controls.Add(this.txt_UserName);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(567, 65);
            this.panel1.TabIndex = 11;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(11, 21);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(41, 12);
            this.label8.TabIndex = 14;
            this.label8.Text = "用户名";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(299, 21);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 12);
            this.label3.TabIndex = 13;
            this.label3.Text = "确认密码";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(161, 21);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(29, 12);
            this.label2.TabIndex = 12;
            this.label2.Text = "密码";
            // 
            // txt_AddPsw2
            // 
            this.txt_AddPsw2.Location = new System.Drawing.Point(358, 17);
            this.txt_AddPsw2.Name = "txt_AddPsw2";
            this.txt_AddPsw2.PasswordChar = '*';
            this.txt_AddPsw2.Size = new System.Drawing.Size(95, 21);
            this.txt_AddPsw2.TabIndex = 11;
            // 
            // txt_AddPsw1
            // 
            this.txt_AddPsw1.Location = new System.Drawing.Point(196, 17);
            this.txt_AddPsw1.Name = "txt_AddPsw1";
            this.txt_AddPsw1.PasswordChar = '*';
            this.txt_AddPsw1.Size = new System.Drawing.Size(95, 21);
            this.txt_AddPsw1.TabIndex = 10;
            // 
            // btn_AddUser
            // 
            this.btn_AddUser.Location = new System.Drawing.Point(457, 13);
            this.btn_AddUser.Name = "btn_AddUser";
            this.btn_AddUser.Size = new System.Drawing.Size(98, 28);
            this.btn_AddUser.TabIndex = 3;
            this.btn_AddUser.Text = "添加用户";
            this.btn_AddUser.UseVisualStyleBackColor = true;
            this.btn_AddUser.Click += new System.EventHandler(this.btn_AddUser_Click);
            // 
            // txt_UserName
            // 
            this.txt_UserName.Location = new System.Drawing.Point(54, 17);
            this.txt_UserName.Name = "txt_UserName";
            this.txt_UserName.Size = new System.Drawing.Size(101, 21);
            this.txt_UserName.TabIndex = 2;
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 65);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.lst_AllUsers);
            this.splitContainer1.Panel1.Controls.Add(this.label1);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.lst_Roles);
            this.splitContainer1.Panel2.Controls.Add(this.label4);
            this.splitContainer1.Panel2.Controls.Add(this.panel2);
            this.splitContainer1.Size = new System.Drawing.Size(567, 345);
            this.splitContainer1.SplitterDistance = 164;
            this.splitContainer1.TabIndex = 12;
            // 
            // lst_AllUsers
            // 
            this.lst_AllUsers.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lst_AllUsers.FormattingEnabled = true;
            this.lst_AllUsers.ItemHeight = 12;
            this.lst_AllUsers.Location = new System.Drawing.Point(0, 14);
            this.lst_AllUsers.Name = "lst_AllUsers";
            this.lst_AllUsers.Size = new System.Drawing.Size(164, 328);
            this.lst_AllUsers.TabIndex = 1;
            this.lst_AllUsers.SelectedIndexChanged += new System.EventHandler(this.lst_AllUsers_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Dock = System.Windows.Forms.DockStyle.Top;
            this.label1.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(0, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(63, 14);
            this.label1.TabIndex = 0;
            this.label1.Text = "用户列表";
            // 
            // lst_Roles
            // 
            this.lst_Roles.CheckOnClick = true;
            this.lst_Roles.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lst_Roles.FormattingEnabled = true;
            this.lst_Roles.Location = new System.Drawing.Point(0, 121);
            this.lst_Roles.Name = "lst_Roles";
            this.lst_Roles.Size = new System.Drawing.Size(399, 212);
            this.lst_Roles.TabIndex = 2;
            this.lst_Roles.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.lst_Roles_ItemCheck);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Dock = System.Windows.Forms.DockStyle.Top;
            this.label4.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(0, 107);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(63, 14);
            this.label4.TabIndex = 1;
            this.label4.Text = "所属角色";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.btn_DelUser);
            this.panel2.Controls.Add(this.lbl_Name);
            this.panel2.Controls.Add(this.btn_ChangePsw);
            this.panel2.Controls.Add(this.label5);
            this.panel2.Controls.Add(this.label6);
            this.panel2.Controls.Add(this.txt_ChgPsw2);
            this.panel2.Controls.Add(this.txt_ChgPsw1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(399, 107);
            this.panel2.TabIndex = 0;
            // 
            // btn_DelUser
            // 
            this.btn_DelUser.Location = new System.Drawing.Point(114, 67);
            this.btn_DelUser.Name = "btn_DelUser";
            this.btn_DelUser.Size = new System.Drawing.Size(98, 28);
            this.btn_DelUser.TabIndex = 21;
            this.btn_DelUser.Text = "删除用户";
            this.btn_DelUser.UseVisualStyleBackColor = true;
            this.btn_DelUser.Click += new System.EventHandler(this.btn_DelUser_Click);
            // 
            // lbl_Name
            // 
            this.lbl_Name.AutoSize = true;
            this.lbl_Name.Font = new System.Drawing.Font("宋体", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lbl_Name.Location = new System.Drawing.Point(40, 36);
            this.lbl_Name.Name = "lbl_Name";
            this.lbl_Name.Size = new System.Drawing.Size(72, 20);
            this.lbl_Name.TabIndex = 19;
            this.lbl_Name.Text = "用户名";
            // 
            // btn_ChangePsw
            // 
            this.btn_ChangePsw.Location = new System.Drawing.Point(286, 67);
            this.btn_ChangePsw.Name = "btn_ChangePsw";
            this.btn_ChangePsw.Size = new System.Drawing.Size(98, 28);
            this.btn_ChangePsw.TabIndex = 18;
            this.btn_ChangePsw.Text = "修改密码";
            this.btn_ChangePsw.UseVisualStyleBackColor = true;
            this.btn_ChangePsw.Click += new System.EventHandler(this.btn_ChangePsw_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(230, 49);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(53, 12);
            this.label5.TabIndex = 17;
            this.label5.Text = "确认密码";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(254, 16);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(29, 12);
            this.label6.TabIndex = 16;
            this.label6.Text = "密码";
            // 
            // txt_ChgPsw2
            // 
            this.txt_ChgPsw2.Location = new System.Drawing.Point(289, 40);
            this.txt_ChgPsw2.Name = "txt_ChgPsw2";
            this.txt_ChgPsw2.PasswordChar = '*';
            this.txt_ChgPsw2.Size = new System.Drawing.Size(95, 21);
            this.txt_ChgPsw2.TabIndex = 15;
            // 
            // txt_ChgPsw1
            // 
            this.txt_ChgPsw1.Location = new System.Drawing.Point(289, 13);
            this.txt_ChgPsw1.Name = "txt_ChgPsw1";
            this.txt_ChgPsw1.PasswordChar = '*';
            this.txt_ChgPsw1.Size = new System.Drawing.Size(95, 21);
            this.txt_ChgPsw1.TabIndex = 14;
            // 
            // 用户管理
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(567, 410);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.panel1);
            this.Name = "用户管理";
            this.Text = "用户管理";
            this.Load += new System.EventHandler(this.用户管理_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel1.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.Panel2.PerformLayout();
            this.splitContainer1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txt_AddPsw2;
        private System.Windows.Forms.TextBox txt_AddPsw1;
        private System.Windows.Forms.Button btn_AddUser;
        private System.Windows.Forms.TextBox txt_UserName;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.ListBox lst_AllUsers;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.CheckedListBox lst_Roles;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btn_ChangePsw;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txt_ChgPsw2;
        private System.Windows.Forms.TextBox txt_ChgPsw1;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label lbl_Name;
        private System.Windows.Forms.Button btn_DelUser;

    }
}