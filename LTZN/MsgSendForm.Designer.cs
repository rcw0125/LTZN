namespace LTZN
{
    partial class MsgSendForm
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
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnSend = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label3 = new System.Windows.Forms.Label();
            this.txtContent = new System.Windows.Forms.TextBox();
            this.btnDelAll = new System.Windows.Forms.Button();
            this.btnDelOne = new System.Windows.Forms.Button();
            this.btnAddAll = new System.Windows.Forms.Button();
            this.btnAddOne = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.lstToUsers = new System.Windows.Forms.ListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.lstAllUser = new System.Windows.Forms.ListBox();
            this.grpTyps = new System.Windows.Forms.GroupBox();
            this.rdoTyp1 = new System.Windows.Forms.RadioButton();
            this.rdoTyp2 = new System.Windows.Forms.RadioButton();
            this.rdoTyp3 = new System.Windows.Forms.RadioButton();
            this.flowLayoutPanel1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.grpTyps.SuspendLayout();
            this.SuspendLayout();
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.Controls.Add(this.btnClose);
            this.flowLayoutPanel1.Controls.Add(this.btnSend);
            this.flowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.flowLayoutPanel1.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(0, 387);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Padding = new System.Windows.Forms.Padding(8, 8, 12, 8);
            this.flowLayoutPanel1.Size = new System.Drawing.Size(732, 49);
            this.flowLayoutPanel1.TabIndex = 11;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(634, 11);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 1;
            this.btnClose.Text = "关闭";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnSend
            // 
            this.btnSend.Location = new System.Drawing.Point(544, 11);
            this.btnSend.Margin = new System.Windows.Forms.Padding(3, 3, 12, 3);
            this.btnSend.Name = "btnSend";
            this.btnSend.Size = new System.Drawing.Size(75, 23);
            this.btnSend.TabIndex = 0;
            this.btnSend.Text = "发送";
            this.btnSend.UseVisualStyleBackColor = true;
            this.btnSend.Click += new System.EventHandler(this.btnSend_Click);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.grpTyps);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.txtContent);
            this.panel1.Controls.Add(this.btnDelAll);
            this.panel1.Controls.Add(this.btnDelOne);
            this.panel1.Controls.Add(this.btnAddAll);
            this.panel1.Controls.Add(this.btnAddOne);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.lstToUsers);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.lstAllUser);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(732, 387);
            this.panel1.TabIndex = 12;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(291, 13);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(29, 12);
            this.label3.TabIndex = 19;
            this.label3.Text = "内容";
            // 
            // txtContent
            // 
            this.txtContent.Location = new System.Drawing.Point(293, 28);
            this.txtContent.Multiline = true;
            this.txtContent.Name = "txtContent";
            this.txtContent.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtContent.Size = new System.Drawing.Size(427, 307);
            this.txtContent.TabIndex = 18;
            // 
            // btnDelAll
            // 
            this.btnDelAll.Location = new System.Drawing.Point(126, 217);
            this.btnDelAll.Name = "btnDelAll";
            this.btnDelAll.Size = new System.Drawing.Size(32, 23);
            this.btnDelAll.TabIndex = 17;
            this.btnDelAll.Text = "<<";
            this.btnDelAll.UseVisualStyleBackColor = true;
            this.btnDelAll.Click += new System.EventHandler(this.btnDelAll_Click);
            // 
            // btnDelOne
            // 
            this.btnDelOne.Location = new System.Drawing.Point(126, 179);
            this.btnDelOne.Name = "btnDelOne";
            this.btnDelOne.Size = new System.Drawing.Size(32, 23);
            this.btnDelOne.TabIndex = 16;
            this.btnDelOne.Text = "<";
            this.btnDelOne.UseVisualStyleBackColor = true;
            this.btnDelOne.Click += new System.EventHandler(this.btnDelOne_Click);
            // 
            // btnAddAll
            // 
            this.btnAddAll.Location = new System.Drawing.Point(126, 139);
            this.btnAddAll.Name = "btnAddAll";
            this.btnAddAll.Size = new System.Drawing.Size(32, 23);
            this.btnAddAll.TabIndex = 15;
            this.btnAddAll.Text = ">>";
            this.btnAddAll.UseVisualStyleBackColor = true;
            this.btnAddAll.Click += new System.EventHandler(this.btnAddAll_Click);
            // 
            // btnAddOne
            // 
            this.btnAddOne.Location = new System.Drawing.Point(126, 98);
            this.btnAddOne.Name = "btnAddOne";
            this.btnAddOne.Size = new System.Drawing.Size(32, 23);
            this.btnAddOne.TabIndex = 14;
            this.btnAddOne.Text = ">";
            this.btnAddOne.UseVisualStyleBackColor = true;
            this.btnAddOne.Click += new System.EventHandler(this.btnAddOne_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(162, 13);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(41, 12);
            this.label2.TabIndex = 13;
            this.label2.Text = "发送到";
            // 
            // lstToUsers
            // 
            this.lstToUsers.FormattingEnabled = true;
            this.lstToUsers.ItemHeight = 12;
            this.lstToUsers.Location = new System.Drawing.Point(164, 28);
            this.lstToUsers.Name = "lstToUsers";
            this.lstToUsers.Size = new System.Drawing.Size(108, 352);
            this.lstToUsers.TabIndex = 12;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 11;
            this.label1.Text = "所有用户";
            // 
            // lstAllUser
            // 
            this.lstAllUser.FormattingEnabled = true;
            this.lstAllUser.ItemHeight = 12;
            this.lstAllUser.Location = new System.Drawing.Point(12, 28);
            this.lstAllUser.Name = "lstAllUser";
            this.lstAllUser.Size = new System.Drawing.Size(108, 352);
            this.lstAllUser.TabIndex = 10;
            // 
            // grpTyps
            // 
            this.grpTyps.Controls.Add(this.rdoTyp3);
            this.grpTyps.Controls.Add(this.rdoTyp2);
            this.grpTyps.Controls.Add(this.rdoTyp1);
            this.grpTyps.Location = new System.Drawing.Point(293, 341);
            this.grpTyps.Name = "grpTyps";
            this.grpTyps.Size = new System.Drawing.Size(427, 33);
            this.grpTyps.TabIndex = 20;
            this.grpTyps.TabStop = false;
            this.grpTyps.Text = "类型";
            // 
            // rdoTyp1
            // 
            this.rdoTyp1.AutoSize = true;
            this.rdoTyp1.Location = new System.Drawing.Point(20, 17);
            this.rdoTyp1.Name = "rdoTyp1";
            this.rdoTyp1.Size = new System.Drawing.Size(71, 16);
            this.rdoTyp1.TabIndex = 0;
            this.rdoTyp1.TabStop = true;
            this.rdoTyp1.Text = "生产指导";
            this.rdoTyp1.UseVisualStyleBackColor = true;
            // 
            // rdoTyp2
            // 
            this.rdoTyp2.AutoSize = true;
            this.rdoTyp2.Location = new System.Drawing.Point(106, 17);
            this.rdoTyp2.Name = "rdoTyp2";
            this.rdoTyp2.Size = new System.Drawing.Size(47, 16);
            this.rdoTyp2.TabIndex = 1;
            this.rdoTyp2.TabStop = true;
            this.rdoTyp2.Text = "提示";
            this.rdoTyp2.UseVisualStyleBackColor = true;
            // 
            // rdoTyp3
            // 
            this.rdoTyp3.AutoSize = true;
            this.rdoTyp3.Location = new System.Drawing.Point(173, 17);
            this.rdoTyp3.Name = "rdoTyp3";
            this.rdoTyp3.Size = new System.Drawing.Size(47, 16);
            this.rdoTyp3.TabIndex = 2;
            this.rdoTyp3.TabStop = true;
            this.rdoTyp3.Text = "通知";
            this.rdoTyp3.UseVisualStyleBackColor = true;
            // 
            // MsgSendForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(732, 436);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.flowLayoutPanel1);
            this.Name = "MsgSendForm";
            this.Text = "发送窗体";
            this.Load += new System.EventHandler(this.MsgSendForm_Load);
            this.flowLayoutPanel1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.grpTyps.ResumeLayout(false);
            this.grpTyps.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtContent;
        private System.Windows.Forms.Button btnDelAll;
        private System.Windows.Forms.Button btnDelOne;
        private System.Windows.Forms.Button btnAddAll;
        private System.Windows.Forms.Button btnAddOne;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ListBox lstToUsers;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox lstAllUser;
        private System.Windows.Forms.Button btnSend;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.GroupBox grpTyps;
        private System.Windows.Forms.RadioButton rdoTyp3;
        private System.Windows.Forms.RadioButton rdoTyp2;
        private System.Windows.Forms.RadioButton rdoTyp1;

    }
}