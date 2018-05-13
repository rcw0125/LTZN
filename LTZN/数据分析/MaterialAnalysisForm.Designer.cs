namespace LTZN.数据分析
{
    partial class MaterialAnalysisForm
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MaterialAnalysisForm));
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStrip_fenlei = new System.Windows.Forms.ToolStripComboBox();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripLabel1 = new System.Windows.Forms.ToolStripLabel();
            this.toolStrip_Luci1 = new System.Windows.Forms.ToolStripTextBox();
            this.sjBegin = new LTZN.Common.ToolStripDateTimePicker();
            this.toolStripLabel2 = new System.Windows.Forms.ToolStripLabel();
            this.toolStrip_Luci2 = new System.Windows.Forms.ToolStripTextBox();
            this.sjEnd = new LTZN.Common.ToolStripDateTimePicker();
            this.toolStripButton1 = new System.Windows.Forms.ToolStripButton();
            this.c1FlexGrid2 = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.materialAnalysisDataSetBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.materialAnalysisDataSet = new LTZN.数据分析.MaterialAnalysisDataSet();
            this.toolStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1FlexGrid2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.materialAnalysisDataSetBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.materialAnalysisDataSet)).BeginInit();
            this.SuspendLayout();
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStrip_fenlei,
            this.toolStripSeparator1,
            this.toolStripLabel1,
            this.toolStrip_Luci1,
            this.sjBegin,
            this.toolStripLabel2,
            this.toolStrip_Luci2,
            this.sjEnd,
            this.toolStripButton1});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(880, 25);
            this.toolStrip1.TabIndex = 5;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStrip_fenlei
            // 
            this.toolStrip_fenlei.Items.AddRange(new object[] {
            "原料",
            "煤粉",
            "焦炭"});
            this.toolStrip_fenlei.Name = "toolStrip_fenlei";
            this.toolStrip_fenlei.Size = new System.Drawing.Size(80, 25);
            this.toolStrip_fenlei.Text = "原料";
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
            // toolStrip_Luci1
            // 
            this.toolStrip_Luci1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.toolStrip_Luci1.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.toolStrip_Luci1.Name = "toolStrip_Luci1";
            this.toolStrip_Luci1.Size = new System.Drawing.Size(120, 25);
            this.toolStrip_Luci1.Visible = false;
            // 
            // sjBegin
            // 
            this.sjBegin.ControlWidth = 150;
            this.sjBegin.CustomFormat = "yyyy年MM月dd日 HH:mm";
            this.sjBegin.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.sjBegin.Name = "sjBegin";
            this.sjBegin.Size = new System.Drawing.Size(150, 22);
            this.sjBegin.Text = "2006年05月24日 16:23";
            this.sjBegin.Value = new System.DateTime(2006, 5, 24, 16, 23, 36, 1);
            // 
            // toolStripLabel2
            // 
            this.toolStripLabel2.Name = "toolStripLabel2";
            this.toolStripLabel2.Size = new System.Drawing.Size(53, 22);
            this.toolStripLabel2.Text = "结束时间";
            // 
            // toolStrip_Luci2
            // 
            this.toolStrip_Luci2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.toolStrip_Luci2.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.toolStrip_Luci2.Name = "toolStrip_Luci2";
            this.toolStrip_Luci2.Size = new System.Drawing.Size(120, 25);
            this.toolStrip_Luci2.Visible = false;
            // 
            // sjEnd
            // 
            this.sjEnd.ControlWidth = 150;
            this.sjEnd.CustomFormat = "yyyy年MM月dd日 HH:mm";
            this.sjEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.sjEnd.Name = "sjEnd";
            this.sjEnd.Size = new System.Drawing.Size(150, 22);
            this.sjEnd.Text = "2006年05月24日 16:23";
            this.sjEnd.Value = new System.DateTime(2006, 5, 24, 16, 23, 36, 1);
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
            // c1FlexGrid2
            // 
            this.c1FlexGrid2.AllowAddNew = true;
            this.c1FlexGrid2.AllowDelete = true;
            this.c1FlexGrid2.AllowDragging = C1.Win.C1FlexGrid.AllowDraggingEnum.None;
            this.c1FlexGrid2.AllowSorting = C1.Win.C1FlexGrid.AllowSortingEnum.None;
            this.c1FlexGrid2.AutoClipboard = true;
            this.c1FlexGrid2.ColumnInfo = resources.GetString("c1FlexGrid2.ColumnInfo");
            this.c1FlexGrid2.DataMember = "DDYL";
            this.c1FlexGrid2.DataSource = this.materialAnalysisDataSetBindingSource;
            this.c1FlexGrid2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.c1FlexGrid2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.c1FlexGrid2.ImeMode = System.Windows.Forms.ImeMode.Alpha;
            this.c1FlexGrid2.KeyActionEnter = C1.Win.C1FlexGrid.KeyActionEnum.MoveAcross;
            this.c1FlexGrid2.Location = new System.Drawing.Point(0, 25);
            this.c1FlexGrid2.Name = "c1FlexGrid2";
            this.c1FlexGrid2.Rows.Count = 1;
            this.c1FlexGrid2.Rows.DefaultSize = 23;
            this.c1FlexGrid2.SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.Row;
            this.c1FlexGrid2.Size = new System.Drawing.Size(880, 348);
            this.c1FlexGrid2.StyleInfo = resources.GetString("c1FlexGrid2.StyleInfo");
            this.c1FlexGrid2.TabIndex = 11;
            this.c1FlexGrid2.Tag = "";
            // 
            // materialAnalysisDataSetBindingSource
            // 
            this.materialAnalysisDataSetBindingSource.DataSource = this.materialAnalysisDataSet;
            this.materialAnalysisDataSetBindingSource.Position = 0;
            // 
            // materialAnalysisDataSet
            // 
            this.materialAnalysisDataSet.DataSetName = "MaterialAnalysisDataSet";
            this.materialAnalysisDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // MaterialAnalysisForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(880, 373);
            this.Controls.Add(this.c1FlexGrid2);
            this.Controls.Add(this.toolStrip1);
            this.Name = "MaterialAnalysisForm";
            this.Text = "原料分析";
            this.Load += new System.EventHandler(this.MaterialAnalysisForm_Load);
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1FlexGrid2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.materialAnalysisDataSetBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.materialAnalysisDataSet)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripComboBox toolStrip_fenlei;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripLabel toolStripLabel1;
        private System.Windows.Forms.ToolStripTextBox toolStrip_Luci1;
        private LTZN.Common.ToolStripDateTimePicker sjBegin;
        private System.Windows.Forms.ToolStripLabel toolStripLabel2;
        private System.Windows.Forms.ToolStripTextBox toolStrip_Luci2;
        private LTZN.Common.ToolStripDateTimePicker sjEnd;
        private System.Windows.Forms.ToolStripButton toolStripButton1;
        private C1.Win.C1FlexGrid.C1FlexGrid c1FlexGrid2;
        private System.Windows.Forms.BindingSource materialAnalysisDataSetBindingSource;
        private MaterialAnalysisDataSet materialAnalysisDataSet;
    }
}