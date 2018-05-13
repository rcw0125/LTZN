namespace LTZN.UserManager
{
    partial class ManGroupForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ManGroupForm));
            this.topPanel = new System.Windows.Forms.Panel();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.txt_AddPsw1 = new System.Windows.Forms.TextBox();
            this.btn_AddUser = new System.Windows.Forms.Button();
            this.txt_UserName = new System.Windows.Forms.TextBox();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.flexGroup = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.c1DockingTab1 = new C1.Win.C1Command.C1DockingTab();
            this.c1DockingTabPage1 = new C1.Win.C1Command.C1DockingTabPage();
            this.flexAuthority = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.c1DockingTabPage2 = new C1.Win.C1Command.C1DockingTabPage();
            this.flexGroupUsers = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.topPanel.SuspendLayout();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.flexGroup)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1DockingTab1)).BeginInit();
            this.c1DockingTab1.SuspendLayout();
            this.c1DockingTabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.flexAuthority)).BeginInit();
            this.c1DockingTabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.flexGroupUsers)).BeginInit();
            this.SuspendLayout();
            // 
            // topPanel
            // 
            this.topPanel.BackgroundImage = global::LTZN.Properties.Resources.duote_08;
            this.topPanel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.topPanel.Controls.Add(this.button1);
            this.topPanel.Controls.Add(this.label1);
            this.topPanel.Controls.Add(this.label8);
            this.topPanel.Controls.Add(this.txt_AddPsw1);
            this.topPanel.Controls.Add(this.btn_AddUser);
            this.topPanel.Controls.Add(this.txt_UserName);
            this.topPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.topPanel.Location = new System.Drawing.Point(0, 0);
            this.topPanel.Name = "topPanel";
            this.topPanel.Size = new System.Drawing.Size(706, 50);
            this.topPanel.TabIndex = 2;
            this.topPanel.Paint += new System.Windows.Forms.PaintEventHandler(this.panel1_Paint);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(495, 17);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(68, 28);
            this.button1.TabIndex = 23;
            this.button1.Text = "保存";
            this.button1.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label1.Location = new System.Drawing.Point(217, 21);
            this.label1.Margin = new System.Windows.Forms.Padding(0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(37, 14);
            this.label1.TabIndex = 22;
            this.label1.Text = "描述";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.BackColor = System.Drawing.Color.Transparent;
            this.label8.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label8.ForeColor = System.Drawing.SystemColors.HotTrack;
            this.label8.Location = new System.Drawing.Point(13, 21);
            this.label8.Margin = new System.Windows.Forms.Padding(0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(67, 14);
            this.label8.TabIndex = 21;
            this.label8.Text = "用户组名";
            this.label8.Click += new System.EventHandler(this.label8_Click);
            // 
            // txt_AddPsw1
            // 
            this.txt_AddPsw1.Location = new System.Drawing.Point(257, 17);
            this.txt_AddPsw1.Name = "txt_AddPsw1";
            this.txt_AddPsw1.PasswordChar = '*';
            this.txt_AddPsw1.Size = new System.Drawing.Size(158, 23);
            this.txt_AddPsw1.TabIndex = 17;
            // 
            // btn_AddUser
            // 
            this.btn_AddUser.Location = new System.Drawing.Point(421, 17);
            this.btn_AddUser.Name = "btn_AddUser";
            this.btn_AddUser.Size = new System.Drawing.Size(68, 28);
            this.btn_AddUser.TabIndex = 16;
            this.btn_AddUser.Text = "添加";
            this.btn_AddUser.UseVisualStyleBackColor = true;
            // 
            // txt_UserName
            // 
            this.txt_UserName.Location = new System.Drawing.Point(83, 17);
            this.txt_UserName.Name = "txt_UserName";
            this.txt_UserName.Size = new System.Drawing.Size(101, 23);
            this.txt_UserName.TabIndex = 15;
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 50);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.flexGroup);
            this.splitContainer1.Panel1.Paint += new System.Windows.Forms.PaintEventHandler(this.splitContainer1_Panel1_Paint);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.c1DockingTab1);
            this.splitContainer1.Size = new System.Drawing.Size(706, 370);
            this.splitContainer1.SplitterDistance = 185;
            this.splitContainer1.SplitterWidth = 5;
            this.splitContainer1.TabIndex = 3;
            // 
            // flexGroup
            // 
            this.flexGroup.ColumnInfo = "3,1,0,0,0,100,Columns:";
            this.flexGroup.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flexGroup.Location = new System.Drawing.Point(0, 0);
            this.flexGroup.Name = "flexGroup";
            this.flexGroup.Rows.Count = 2;
            this.flexGroup.Rows.DefaultSize = 20;
            this.flexGroup.Size = new System.Drawing.Size(706, 185);
            this.flexGroup.StyleInfo = resources.GetString("flexGroup.StyleInfo");
            this.flexGroup.TabIndex = 1;
            this.flexGroup.Click += new System.EventHandler(this.c1FlexGrid1_Click_1);
            // 
            // c1DockingTab1
            // 
            this.c1DockingTab1.Controls.Add(this.c1DockingTabPage1);
            this.c1DockingTab1.Controls.Add(this.c1DockingTabPage2);
            this.c1DockingTab1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.c1DockingTab1.Location = new System.Drawing.Point(0, 0);
            this.c1DockingTab1.Name = "c1DockingTab1";
            this.c1DockingTab1.SelectedIndex = 1;
            this.c1DockingTab1.Size = new System.Drawing.Size(706, 180);
            this.c1DockingTab1.TabIndex = 0;
            this.c1DockingTab1.TabsSpacing = 0;
            this.c1DockingTab1.TabStyle = C1.Win.C1Command.TabStyleEnum.WindowsXP;
            this.c1DockingTab1.VisualStyle = C1.Win.C1Command.VisualStyle.Custom;
            this.c1DockingTab1.VisualStyleBase = C1.Win.C1Command.VisualStyle.WindowsXP;
            // 
            // c1DockingTabPage1
            // 
            this.c1DockingTabPage1.Controls.Add(this.flexAuthority);
            this.c1DockingTabPage1.Location = new System.Drawing.Point(2, 29);
            this.c1DockingTabPage1.Name = "c1DockingTabPage1";
            this.c1DockingTabPage1.Size = new System.Drawing.Size(700, 147);
            this.c1DockingTabPage1.TabIndex = 2;
            this.c1DockingTabPage1.Text = "权限";
            // 
            // flexAuthority
            // 
            this.flexAuthority.BorderStyle = C1.Win.C1FlexGrid.Util.BaseControls.BorderStyleEnum.XpThemes;
            this.flexAuthority.ColumnInfo = "3,1,0,0,0,100,Columns:";
            this.flexAuthority.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flexAuthority.Location = new System.Drawing.Point(0, 0);
            this.flexAuthority.Name = "flexAuthority";
            this.flexAuthority.Rows.Count = 2;
            this.flexAuthority.Rows.DefaultSize = 20;
            this.flexAuthority.Size = new System.Drawing.Size(700, 147);
            this.flexAuthority.StyleInfo = resources.GetString("flexAuthority.StyleInfo");
            this.flexAuthority.TabIndex = 3;
            this.flexAuthority.Click += new System.EventHandler(this.c1FlexGrid2_Click_2);
            // 
            // c1DockingTabPage2
            // 
            this.c1DockingTabPage2.Controls.Add(this.flexGroupUsers);
            this.c1DockingTabPage2.Location = new System.Drawing.Point(2, 29);
            this.c1DockingTabPage2.Name = "c1DockingTabPage2";
            this.c1DockingTabPage2.Size = new System.Drawing.Size(700, 147);
            this.c1DockingTabPage2.TabIndex = 1;
            this.c1DockingTabPage2.Text = "用户";
            // 
            // flexGroupUsers
            // 
            this.flexGroupUsers.BorderStyle = C1.Win.C1FlexGrid.Util.BaseControls.BorderStyleEnum.XpThemes;
            this.flexGroupUsers.ColumnInfo = "10,1,0,0,0,100,Columns:";
            this.flexGroupUsers.Dock = System.Windows.Forms.DockStyle.Fill;
            this.flexGroupUsers.Location = new System.Drawing.Point(0, 0);
            this.flexGroupUsers.Name = "flexGroupUsers";
            this.flexGroupUsers.Rows.Count = 2;
            this.flexGroupUsers.Rows.DefaultSize = 20;
            this.flexGroupUsers.Size = new System.Drawing.Size(700, 147);
            this.flexGroupUsers.StyleInfo = resources.GetString("flexGroupUsers.StyleInfo");
            this.flexGroupUsers.TabIndex = 4;
            // 
            // ManGroupForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 14F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(706, 420);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.topPanel);
            this.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Name = "ManGroupForm";
            this.Text = "用户组管理";
            this.Load += new System.EventHandler(this.ManGroupForm_Load);
            this.topPanel.ResumeLayout(false);
            this.topPanel.PerformLayout();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.flexGroup)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1DockingTab1)).EndInit();
            this.c1DockingTab1.ResumeLayout(false);
            this.c1DockingTabPage1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.flexAuthority)).EndInit();
            this.c1DockingTabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.flexGroupUsers)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel topPanel;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private C1.Win.C1FlexGrid.C1FlexGrid flexGroup;
        private C1.Win.C1Command.C1DockingTab c1DockingTab1;
        private C1.Win.C1Command.C1DockingTabPage c1DockingTabPage1;
        private C1.Win.C1FlexGrid.C1FlexGrid flexAuthority;
        private C1.Win.C1Command.C1DockingTabPage c1DockingTabPage2;
        private C1.Win.C1FlexGrid.C1FlexGrid flexGroupUsers;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txt_AddPsw1;
        private System.Windows.Forms.Button btn_AddUser;
        private System.Windows.Forms.TextBox txt_UserName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button1;
    }
}