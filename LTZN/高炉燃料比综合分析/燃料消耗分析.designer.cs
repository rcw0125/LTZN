namespace LTZN.高炉燃料比综合分析
{
    partial class 燃料消耗分析
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(燃料消耗分析));
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.gaolu = new System.Windows.Forms.ToolStripComboBox();
            this.toolStripLabel3 = new System.Windows.Forms.ToolStripLabel();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripLabel1 = new System.Windows.Forms.ToolStripLabel();
            this.sjBegin = new LTZN.Common.ToolStripDateTimePicker();
            this.toolStripLabel2 = new System.Windows.Forms.ToolStripLabel();
            this.sjEnd = new LTZN.Common.ToolStripDateTimePicker();
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.c1Chart1 = new C1.Win.C1Chart.C1Chart();
            this.c1Chart2 = new C1.Win.C1Chart.C1Chart();
            this.toolStrip1.SuspendLayout();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1Chart1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1Chart2)).BeginInit();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.gaolu,
            this.toolStripLabel3,
            this.toolStripSeparator1,
            this.toolStripLabel1,
            this.sjBegin,
            this.toolStripLabel2,
            this.sjEnd,
            this.toolStripButton1});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(741, 25);
            this.toolStrip1.TabIndex = 4;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // gaolu
            // 
            this.gaolu.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5",
            "6"});
            this.gaolu.Name = "gaolu";
            this.gaolu.Size = new System.Drawing.Size(75, 25);
            this.gaolu.Text = "1";
            // 
            // toolStripLabel3
            // 
            this.toolStripLabel3.Name = "toolStripLabel3";
            this.toolStripLabel3.Size = new System.Drawing.Size(29, 22);
            this.toolStripLabel3.Text = "高炉";
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 25);
            // 
            // toolStripLabel1
            // 
            this.toolStripLabel1.Name = "toolStripLabel1";
            this.toolStripLabel1.Size = new System.Drawing.Size(53, 22);
            this.toolStripLabel1.Text = "开始时间";
            // 
            // sjBegin
            // 
            this.sjBegin.ControlWidth = 150;
            this.sjBegin.CustomFormat = "yyyy年MM月dd日";
            this.sjBegin.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.sjBegin.Name = "sjBegin";
            this.sjBegin.Size = new System.Drawing.Size(150, 22);
            this.sjBegin.Text = "2013年03月01日";
            this.sjBegin.Value = new System.DateTime(2013, 3, 1, 0, 0, 0, 0);
            // 
            // toolStripLabel2
            // 
            this.toolStripLabel2.Name = "toolStripLabel2";
            this.toolStripLabel2.Size = new System.Drawing.Size(53, 22);
            this.toolStripLabel2.Text = "结束时间";
            // 
            // sjEnd
            // 
            this.sjEnd.ControlWidth = 150;
            this.sjEnd.CustomFormat = "yyyy年MM月dd日";
            this.sjEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.sjEnd.Name = "sjEnd";
            this.sjEnd.Size = new System.Drawing.Size(150, 22);
            this.sjEnd.Text = "2013年03月03日";
            this.sjEnd.Value = new System.DateTime(2013, 3, 3, 0, 0, 0, 0);
            // 
            // toolStripButton1
            // 
            this.toolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton1.Image")));
            this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton1.Name = "toolStripButton1";
            this.toolStripButton1.Size = new System.Drawing.Size(33, 22);
            this.toolStripButton1.Text = "查询";
            this.toolStripButton1.Click += new System.EventHandler(this.toolStripButton1_Click);
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 25);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.c1Chart1);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.c1Chart2);
            this.splitContainer1.Panel2Collapsed = true;
            this.splitContainer1.Size = new System.Drawing.Size(741, 332);
            this.splitContainer1.SplitterDistance = 165;
            this.splitContainer1.TabIndex = 5;
            // 
            // c1Chart1
            // 
            this.c1Chart1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.c1Chart1.Location = new System.Drawing.Point(0, 0);
            this.c1Chart1.Name = "c1Chart1";
            this.c1Chart1.PropBag = resources.GetString("c1Chart1.PropBag");
            this.c1Chart1.Size = new System.Drawing.Size(741, 332);
            this.c1Chart1.TabIndex = 5;
            // 
            // c1Chart2
            // 
            this.c1Chart2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.c1Chart2.Location = new System.Drawing.Point(0, 0);
            this.c1Chart2.Name = "c1Chart2";
            this.c1Chart2.PropBag = resources.GetString("c1Chart2.PropBag");
            this.c1Chart2.Size = new System.Drawing.Size(150, 46);
            this.c1Chart2.TabIndex = 5;
            // 
            // 燃料消耗分析
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(741, 357);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.toolStrip1);
            this.Name = "燃料消耗分析";
            this.Text = "燃料消耗分析";
            this.Load += new System.EventHandler(this.燃料消耗分析_Load);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.c1Chart1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1Chart2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripLabel toolStripLabel1;
        private LTZN.Common.ToolStripDateTimePicker sjBegin;
        private System.Windows.Forms.ToolStripLabel toolStripLabel2;
        private LTZN.Common.ToolStripDateTimePicker sjEnd;
        private System.Windows.Forms.ToolStripButton toolStripButton1;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private C1.Win.C1Chart.C1Chart c1Chart1;
        private C1.Win.C1Chart.C1Chart c1Chart2;
        private System.Windows.Forms.ToolStripLabel toolStripLabel3;
        private System.Windows.Forms.ToolStripComboBox gaolu;
    }
}