namespace LTZN.高炉燃料比综合分析
{
    partial class 数据汇总
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(数据汇总));
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.gaolu = new System.Windows.Forms.ToolStripComboBox();
            this.toolStripLabel3 = new System.Windows.Forms.ToolStripLabel();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripLabel1 = new System.Windows.Forms.ToolStripLabel();
            this.sjBegin = new LTZN.Common.ToolStripDateTimePicker();
            this.toolStripLabel2 = new System.Windows.Forms.ToolStripLabel();
            this.sjEnd = new LTZN.Common.ToolStripDateTimePicker();
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.保存SToolStripButton = new System.Windows.Forms.ToolStripButton();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.c1FlexGrid_tiezha = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.c1Chart2 = new C1.Win.C1Chart.C1Chart();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.toolStrip1.SuspendLayout();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1FlexGrid_tiezha)).BeginInit();
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
            this.toolStripButton1,
            this.保存SToolStripButton});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(741, 26);
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
            this.gaolu.Size = new System.Drawing.Size(75, 26);
            this.gaolu.Text = "1";
            // 
            // toolStripLabel3
            // 
            this.toolStripLabel3.Name = "toolStripLabel3";
            this.toolStripLabel3.Size = new System.Drawing.Size(32, 23);
            this.toolStripLabel3.Text = "高炉";
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 26);
            // 
            // toolStripLabel1
            // 
            this.toolStripLabel1.Name = "toolStripLabel1";
            this.toolStripLabel1.Size = new System.Drawing.Size(56, 23);
            this.toolStripLabel1.Text = "开始时间";
            // 
            // sjBegin
            // 
            this.sjBegin.ControlWidth = 150;
            this.sjBegin.CustomFormat = "yyyy年MM月dd日";
            this.sjBegin.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.sjBegin.Name = "sjBegin";
            this.sjBegin.Size = new System.Drawing.Size(150, 23);
            this.sjBegin.Text = "2013年03月01日";
            this.sjBegin.Value = new System.DateTime(2013, 3, 1, 0, 0, 0, 0);
            // 
            // toolStripLabel2
            // 
            this.toolStripLabel2.Name = "toolStripLabel2";
            this.toolStripLabel2.Size = new System.Drawing.Size(56, 23);
            this.toolStripLabel2.Text = "结束时间";
            // 
            // sjEnd
            // 
            this.sjEnd.ControlWidth = 150;
            this.sjEnd.CustomFormat = "yyyy年MM月dd日";
            this.sjEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.sjEnd.Name = "sjEnd";
            this.sjEnd.Size = new System.Drawing.Size(150, 23);
            this.sjEnd.Text = "2013年03月03日";
            this.sjEnd.Value = new System.DateTime(2013, 3, 3, 0, 0, 0, 0);
            // 
            // toolStripButton1
            // 
            this.toolStripButton1.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripButton1.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButton1.Image")));
            this.toolStripButton1.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButton1.Name = "toolStripButton1";
            this.toolStripButton1.Size = new System.Drawing.Size(36, 23);
            this.toolStripButton1.Text = "查询";
            this.toolStripButton1.Click += new System.EventHandler(this.toolStripButton1_Click);
            // 
            // 保存SToolStripButton
            // 
            this.保存SToolStripButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.保存SToolStripButton.Image = ((System.Drawing.Image)(resources.GetObject("保存SToolStripButton.Image")));
            this.保存SToolStripButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.保存SToolStripButton.Name = "保存SToolStripButton";
            this.保存SToolStripButton.Size = new System.Drawing.Size(23, 23);
            this.保存SToolStripButton.Text = "保存(&S)";
            this.保存SToolStripButton.Click += new System.EventHandler(this.保存SToolStripButton_Click);
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 26);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.c1FlexGrid_tiezha);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.c1Chart2);
            this.splitContainer1.Panel2Collapsed = true;
            this.splitContainer1.Size = new System.Drawing.Size(741, 331);
            this.splitContainer1.SplitterDistance = 165;
            this.splitContainer1.TabIndex = 5;
            // 
            // c1FlexGrid_tiezha
            // 
            this.c1FlexGrid_tiezha.AllowDragging = C1.Win.C1FlexGrid.AllowDraggingEnum.None;
            this.c1FlexGrid_tiezha.AllowEditing = false;
            this.c1FlexGrid_tiezha.AllowSorting = C1.Win.C1FlexGrid.AllowSortingEnum.None;
            this.c1FlexGrid_tiezha.AutoClipboard = true;
            this.c1FlexGrid_tiezha.AutoGenerateColumns = false;
            this.c1FlexGrid_tiezha.ColumnInfo = resources.GetString("c1FlexGrid_tiezha.ColumnInfo");
            this.c1FlexGrid_tiezha.Cursor = System.Windows.Forms.Cursors.Default;
            this.c1FlexGrid_tiezha.Dock = System.Windows.Forms.DockStyle.Fill;
            this.c1FlexGrid_tiezha.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.c1FlexGrid_tiezha.KeyActionEnter = C1.Win.C1FlexGrid.KeyActionEnum.MoveAcross;
            this.c1FlexGrid_tiezha.Location = new System.Drawing.Point(0, 0);
            this.c1FlexGrid_tiezha.Name = "c1FlexGrid_tiezha";
            this.c1FlexGrid_tiezha.Rows.Count = 13;
            this.c1FlexGrid_tiezha.Rows.DefaultSize = 23;
            this.c1FlexGrid_tiezha.SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.Row;
            this.c1FlexGrid_tiezha.Size = new System.Drawing.Size(741, 331);
            this.c1FlexGrid_tiezha.StyleInfo = resources.GetString("c1FlexGrid_tiezha.StyleInfo");
            this.c1FlexGrid_tiezha.TabIndex = 16;
            this.c1FlexGrid_tiezha.Tag = "";
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
            // saveFileDialog1
            // 
            this.saveFileDialog1.DefaultExt = "xls";
            this.saveFileDialog1.Filter = "Excel 文件|*.xls";
            this.saveFileDialog1.FileOk += new System.ComponentModel.CancelEventHandler(this.saveFileDialog1_FileOk);
            // 
            // 数据汇总
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(741, 357);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.toolStrip1);
            this.Name = "数据汇总";
            this.Text = "数据汇总";
            this.Load += new System.EventHandler(this.数据汇总_Load);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.c1FlexGrid_tiezha)).EndInit();
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
        private System.Windows.Forms.SplitContainer splitContainer1;
        private C1.Win.C1Chart.C1Chart c1Chart2;
        private System.Windows.Forms.ToolStripLabel toolStripLabel3;
        private System.Windows.Forms.ToolStripComboBox gaolu;
        private C1.Win.C1FlexGrid.C1FlexGrid c1FlexGrid_tiezha;
        private System.Windows.Forms.ToolStripButton toolStripButton1;
        private System.Windows.Forms.ToolStripButton 保存SToolStripButton;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
    }
}