namespace LTZN
{
    partial class MsgContentForm
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtContent = new System.Windows.Forms.TextBox();
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.btnClose = new System.Windows.Forms.Button();
            this.btnOk = new System.Windows.Forms.Button();
            this.lSendUser = new System.Windows.Forms.Label();
            this.lSendTime = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.lMsgTyp = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.txtHuifu = new System.Windows.Forms.TextBox();
            this.flowLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(36, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "发送人：";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(36, 60);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 1;
            this.label2.Text = "发送时间：";
            // 
            // txtContent
            // 
            this.txtContent.BackColor = System.Drawing.SystemColors.Info;
            this.txtContent.Location = new System.Drawing.Point(12, 91);
            this.txtContent.Multiline = true;
            this.txtContent.Name = "txtContent";
            this.txtContent.ReadOnly = true;
            this.txtContent.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtContent.Size = new System.Drawing.Size(457, 157);
            this.txtContent.TabIndex = 2;
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.Controls.Add(this.btnClose);
            this.flowLayoutPanel1.Controls.Add(this.btnOk);
            this.flowLayoutPanel1.Controls.Add(this.button1);
            this.flowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.flowLayoutPanel1.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(0, 368);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Padding = new System.Windows.Forms.Padding(8, 8, 12, 8);
            this.flowLayoutPanel1.Size = new System.Drawing.Size(485, 46);
            this.flowLayoutPanel1.TabIndex = 3;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(387, 11);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 0;
            this.btnClose.Text = "稍后处理";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnOk
            // 
            this.btnOk.Location = new System.Drawing.Point(306, 11);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(75, 23);
            this.btnOk.TabIndex = 1;
            this.btnOk.Text = "处理";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // lSendUser
            // 
            this.lSendUser.AutoSize = true;
            this.lSendUser.Location = new System.Drawing.Point(112, 21);
            this.lSendUser.Name = "lSendUser";
            this.lSendUser.Size = new System.Drawing.Size(0, 12);
            this.lSendUser.TabIndex = 4;
            // 
            // lSendTime
            // 
            this.lSendTime.AutoSize = true;
            this.lSendTime.Location = new System.Drawing.Point(112, 60);
            this.lSendTime.Name = "lSendTime";
            this.lSendTime.Size = new System.Drawing.Size(0, 12);
            this.lSendTime.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(239, 21);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(65, 12);
            this.label3.TabIndex = 6;
            this.label3.Text = "消息类型：";
            // 
            // lMsgTyp
            // 
            this.lMsgTyp.AutoSize = true;
            this.lMsgTyp.Location = new System.Drawing.Point(312, 21);
            this.lMsgTyp.Name = "lMsgTyp";
            this.lMsgTyp.Size = new System.Drawing.Size(0, 12);
            this.lMsgTyp.TabIndex = 7;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(225, 11);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 2;
            this.button1.Text = "回复";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtHuifu
            // 
            this.txtHuifu.Location = new System.Drawing.Point(12, 265);
            this.txtHuifu.Multiline = true;
            this.txtHuifu.Name = "txtHuifu";
            this.txtHuifu.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.txtHuifu.Size = new System.Drawing.Size(457, 97);
            this.txtHuifu.TabIndex = 8;
            // 
            // MsgContentForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(485, 414);
            this.Controls.Add(this.txtHuifu);
            this.Controls.Add(this.lMsgTyp);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.lSendTime);
            this.Controls.Add(this.lSendUser);
            this.Controls.Add(this.flowLayoutPanel1);
            this.Controls.Add(this.txtContent);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "MsgContentForm";
            this.Text = "消息内容";
            this.Load += new System.EventHandler(this.MsgContentForm_Load);
            this.flowLayoutPanel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtContent;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Label lSendUser;
        private System.Windows.Forms.Label lSendTime;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lMsgTyp;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox txtHuifu;
    }
}