namespace LTZN
{
    partial class MsgUnDealForm
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
            this.listView1 = new System.Windows.Forms.ListView();
            this.lhSendUser = new System.Windows.Forms.ColumnHeader();
            this.lhSendTime = new System.Windows.Forms.ColumnHeader();
            this.lhTyp = new System.Windows.Forms.ColumnHeader();
            this.SuspendLayout();
            // 
            // listView1
            // 
            this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.lhSendUser,
            this.lhSendTime,
            this.lhTyp});
            this.listView1.Location = new System.Drawing.Point(3, 3);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(558, 370);
            this.listView1.TabIndex = 0;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            this.listView1.DoubleClick += new System.EventHandler(this.listView1_DoubleClick);
            // 
            // lhSendUser
            // 
            this.lhSendUser.Text = "发送人";
            this.lhSendUser.Width = 90;
            // 
            // lhSendTime
            // 
            this.lhSendTime.Text = "发送时间";
            this.lhSendTime.Width = 100;
            // 
            // lhTyp
            // 
            this.lhTyp.Text = "消息类型";
            this.lhTyp.Width = 100;
            // 
            // MsgUnDealForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(573, 385);
            this.Controls.Add(this.listView1);
            this.Name = "MsgUnDealForm";
            this.Text = "未处理的消息";
            this.Load += new System.EventHandler(this.MsgUnDealForm_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.ColumnHeader lhSendUser;
        private System.Windows.Forms.ColumnHeader lhSendTime;
        private System.Windows.Forms.ColumnHeader lhTyp;
    }
}