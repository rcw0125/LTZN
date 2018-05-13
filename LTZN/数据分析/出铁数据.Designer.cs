namespace LTZN.数据分析
{
    partial class 出铁数据
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(出铁数据));
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStrip_gaolu = new System.Windows.Forms.ToolStripComboBox();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.r1 = new LTZN.Common.ToolStripRadioButtonPicker();
            this.r2 = new LTZN.Common.ToolStripRadioButtonPicker();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripLabel1 = new System.Windows.Forms.ToolStripLabel();
            this.toolStrip_Luci1 = new System.Windows.Forms.ToolStripTextBox();
            this.sjBegin = new LTZN.Common.ToolStripDateTimePicker();
            this.toolStripLabel2 = new System.Windows.Forms.ToolStripLabel();
            this.toolStrip_Luci2 = new System.Windows.Forms.ToolStripTextBox();
            this.sjEnd = new LTZN.Common.ToolStripDateTimePicker();
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.label1 = new System.Windows.Forms.Label();
            this.c1FlexGrid3 = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.dDLUCIBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.chuTieDataSet = new LTZN.数据分析.ChuTieDataSet();
            this.dDLUCITableAdapter = new LTZN.数据分析.ChuTieDataSetTableAdapters.DDLUCITableAdapter();
            this.toolStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1FlexGrid3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dDLUCIBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.chuTieDataSet)).BeginInit();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStrip_gaolu,
            this.toolStripSeparator1,
            this.r1,
            this.r2,
            this.toolStripSeparator3,
            this.toolStripLabel1,
            this.toolStrip_Luci1,
            this.sjBegin,
            this.toolStripLabel2,
            this.toolStrip_Luci2,
            this.sjEnd,
            this.toolStripButton1});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(890, 26);
            this.toolStrip1.TabIndex = 4;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStrip_gaolu
            // 
            this.toolStrip_gaolu.Items.AddRange(new object[] {
            "1高炉",
            "2高炉",
            "3高炉",
            "4高炉",
            "5高炉",
            "6高炉",
            "全部"});
            this.toolStrip_gaolu.Name = "toolStrip_gaolu";
            this.toolStrip_gaolu.Size = new System.Drawing.Size(120, 26);
            this.toolStrip_gaolu.Text = "1高炉";
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 26);
            // 
            // r1
            // 
            this.r1.Font = new System.Drawing.Font("宋体", 9F);
            this.r1.Name = "r1";
            this.r1.RadioCheckAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.r1.RadioChecked = true;
            this.r1.RadioFont = new System.Drawing.Font("宋体", 9F);
            this.r1.RadioHeight = 23;
            this.r1.RadioText = "按时间";
            this.r1.RadioTextImageRelation = System.Windows.Forms.TextImageRelation.Overlay;
            this.r1.RadioWidth = 59;
            this.r1.Size = new System.Drawing.Size(59, 23);
            this.r1.Text = "按时间";
            this.r1.CheckedChanged += new System.EventHandler(this.r1_CheckedChanged);
            // 
            // r2
            // 
            this.r2.Font = new System.Drawing.Font("宋体", 9F);
            this.r2.Name = "r2";
            this.r2.RadioCheckAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.r2.RadioChecked = false;
            this.r2.RadioFont = new System.Drawing.Font("宋体", 9F);
            this.r2.RadioHeight = 23;
            this.r2.RadioText = "按炉次";
            this.r2.RadioTextImageRelation = System.Windows.Forms.TextImageRelation.Overlay;
            this.r2.RadioWidth = 59;
            this.r2.Size = new System.Drawing.Size(59, 23);
            this.r2.Text = "按炉次";
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(6, 26);
            // 
            // toolStripLabel1
            // 
            this.toolStripLabel1.Name = "toolStripLabel1";
            this.toolStripLabel1.Size = new System.Drawing.Size(56, 23);
            this.toolStripLabel1.Text = "开始时间";
            // 
            // toolStrip_Luci1
            // 
            this.toolStrip_Luci1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.toolStrip_Luci1.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.toolStrip_Luci1.Name = "toolStrip_Luci1";
            this.toolStrip_Luci1.Size = new System.Drawing.Size(120, 26);
            this.toolStrip_Luci1.Visible = false;
            this.toolStrip_Luci1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.toolStrip_Luci1_KeyDown);
            // 
            // sjBegin
            // 
            this.sjBegin.ControlWidth = 150;
            this.sjBegin.CustomFormat = "yyyy年MM月dd日 HH:mm";
            this.sjBegin.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.sjBegin.Name = "sjBegin";
            this.sjBegin.Size = new System.Drawing.Size(150, 23);
            this.sjBegin.Text = "2006年05月24日 16:23";
            this.sjBegin.Value = new System.DateTime(2006, 5, 24, 16, 23, 36, 1);
            // 
            // toolStripLabel2
            // 
            this.toolStripLabel2.Name = "toolStripLabel2";
            this.toolStripLabel2.Size = new System.Drawing.Size(56, 23);
            this.toolStripLabel2.Text = "结束时间";
            // 
            // toolStrip_Luci2
            // 
            this.toolStrip_Luci2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.toolStrip_Luci2.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.toolStrip_Luci2.Name = "toolStrip_Luci2";
            this.toolStrip_Luci2.Size = new System.Drawing.Size(120, 26);
            this.toolStrip_Luci2.Visible = false;
            this.toolStrip_Luci2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.toolStrip_Luci2_KeyDown);
            // 
            // sjEnd
            // 
            this.sjEnd.ControlWidth = 150;
            this.sjEnd.CustomFormat = "yyyy年MM月dd日 HH:mm";
            this.sjEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.sjEnd.Name = "sjEnd";
            this.sjEnd.Size = new System.Drawing.Size(150, 23);
            this.sjEnd.Text = "2006年05月24日 16:23";
            this.sjEnd.Value = new System.DateTime(2006, 5, 24, 16, 23, 36, 1);
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
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(700, 60);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 12);
            this.label1.TabIndex = 6;
            // 
            // c1FlexGrid3
            // 
            this.c1FlexGrid3.AllowDragging = C1.Win.C1FlexGrid.AllowDraggingEnum.None;
            this.c1FlexGrid3.AllowEditing = false;
            this.c1FlexGrid3.AllowSorting = C1.Win.C1FlexGrid.AllowSortingEnum.None;
            this.c1FlexGrid3.AutoClipboard = true;
            this.c1FlexGrid3.ColumnInfo = resources.GetString("c1FlexGrid3.ColumnInfo");
            this.c1FlexGrid3.DataSource = this.dDLUCIBindingSource;
            this.c1FlexGrid3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.c1FlexGrid3.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.c1FlexGrid3.ImeMode = System.Windows.Forms.ImeMode.Alpha;
            this.c1FlexGrid3.KeyActionEnter = C1.Win.C1FlexGrid.KeyActionEnum.MoveAcross;
            this.c1FlexGrid3.Location = new System.Drawing.Point(0, 26);
            this.c1FlexGrid3.Name = "c1FlexGrid3";
            this.c1FlexGrid3.Rows.Count = 1;
            this.c1FlexGrid3.Rows.DefaultSize = 23;
            this.c1FlexGrid3.SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.Row;
            this.c1FlexGrid3.Size = new System.Drawing.Size(890, 396);
            this.c1FlexGrid3.StyleInfo = resources.GetString("c1FlexGrid3.StyleInfo");
            this.c1FlexGrid3.TabIndex = 12;
            this.c1FlexGrid3.Tag = "";
            // 
            // dDLUCIBindingSource
            // 
            this.dDLUCIBindingSource.DataMember = "DDLUCI";
            this.dDLUCIBindingSource.DataSource = this.chuTieDataSet;
            // 
            // chuTieDataSet
            // 
            this.chuTieDataSet.DataSetName = "ChuTieDataSet";
            this.chuTieDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // dDLUCITableAdapter
            // 
            this.dDLUCITableAdapter.ClearBeforeFill = true;
            // 
            // 出铁数据
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(890, 422);
            this.Controls.Add(this.c1FlexGrid3);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.toolStrip1);
            this.Name = "出铁数据";
            this.Text = "出铁数据";
            this.Load += new System.EventHandler(this.出铁数据_Load);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1FlexGrid3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dDLUCIBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.chuTieDataSet)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripComboBox toolStrip_gaolu;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private LTZN.Common.ToolStripRadioButtonPicker r1;
        private LTZN.Common.ToolStripRadioButtonPicker r2;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
        private System.Windows.Forms.ToolStripLabel toolStripLabel1;
        private System.Windows.Forms.ToolStripTextBox toolStrip_Luci1;
        private LTZN.Common.ToolStripDateTimePicker sjBegin;
        private System.Windows.Forms.ToolStripLabel toolStripLabel2;
        private System.Windows.Forms.ToolStripTextBox toolStrip_Luci2;
        private LTZN.Common.ToolStripDateTimePicker sjEnd;
        private System.Windows.Forms.ToolStripButton toolStripButton1;
        private System.Windows.Forms.Label label1;
        private C1.Win.C1FlexGrid.C1FlexGrid c1FlexGrid3;
        private ChuTieDataSet chuTieDataSet;
        private System.Windows.Forms.BindingSource dDLUCIBindingSource;
        private LTZN.数据分析.ChuTieDataSetTableAdapters.DDLUCITableAdapter dDLUCITableAdapter;
    }
}