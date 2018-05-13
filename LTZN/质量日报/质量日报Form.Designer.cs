namespace LTZN.质量日报
{
    partial class 质量日报Form
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
            Microsoft.Reporting.WinForms.ReportDataSource reportDataSource1 = new Microsoft.Reporting.WinForms.ReportDataSource();
            Microsoft.Reporting.WinForms.ReportDataSource reportDataSource2 = new Microsoft.Reporting.WinForms.ReportDataSource();
            Microsoft.Reporting.WinForms.ReportDataSource reportDataSource3 = new Microsoft.Reporting.WinForms.ReportDataSource();
            Microsoft.Reporting.WinForms.ReportDataSource reportDataSource4 = new Microsoft.Reporting.WinForms.ReportDataSource();
            Microsoft.Reporting.WinForms.ReportDataSource reportDataSource5 = new Microsoft.Reporting.WinForms.ReportDataSource();
            Microsoft.Reporting.WinForms.ReportDataSource reportDataSource6 = new Microsoft.Reporting.WinForms.ReportDataSource();
            Microsoft.Reporting.WinForms.ReportDataSource reportDataSource7 = new Microsoft.Reporting.WinForms.ReportDataSource();
            Microsoft.Reporting.WinForms.ReportDataSource reportDataSource8 = new Microsoft.Reporting.WinForms.ReportDataSource();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(质量日报Form));
            this.DDYLBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.dDJSBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.dDJS = new LTZN.质量日报.DDJS();
            this.燃料指标内容项BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.reportViewer1 = new Microsoft.Reporting.WinForms.ReportViewer();
            this.webBrowser1 = new System.Windows.Forms.WebBrowser();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripLabel2 = new System.Windows.Forms.ToolStripLabel();
            this.riqiControl = new LTZN.Common.ToolStripDateTimePicker();
            this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
            this.btnQuery = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator6 = new System.Windows.Forms.ToolStripSeparator();
            this.btnSave = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.btnModel = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.btnPrint = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.btnExcel = new System.Windows.Forms.ToolStripButton();
            this.DDYLTableAdapter = new LTZN.质量日报.质量日报数据集TableAdapters.DDYLTableAdapter();
            this.dDJSTableAdapter = new LTZN.质量日报.DDJSTableAdapters.DDJSTableAdapter();
            this.焦炭指标内容项BindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.dDTTY = new LTZN.质量日报.DDTTY();
            this.dDTTYBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.dDTTYTableAdapter = new LTZN.质量日报.DDTTYTableAdapters.DDTTYTableAdapter();
            ((System.ComponentModel.ISupportInitialize)(this.DDYLBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dDJSBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dDJS)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.燃料指标内容项BindingSource)).BeginInit();
            this.toolStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.焦炭指标内容项BindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dDTTY)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dDTTYBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // dDJSBindingSource
            // 
            this.dDJSBindingSource.DataMember = "DDJS";
            this.dDJSBindingSource.DataSource = this.dDJS;
            // 
            // dDJS
            // 
            this.dDJS.DataSetName = "DDJS";
            this.dDJS.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // 燃料指标内容项BindingSource
            // 
            this.燃料指标内容项BindingSource.DataSource = typeof(LTZN.质量日报.燃料指标内容项);
            // 
            // reportViewer1
            // 
            this.reportViewer1.Dock = System.Windows.Forms.DockStyle.Fill;
            reportDataSource1.Name = "技术日报内容";
            reportDataSource1.Value = this.DDYLBindingSource;
            reportDataSource2.Name = "质量日报数据集_DDYL";
            reportDataSource2.Value = this.DDYLBindingSource;
            reportDataSource3.Name = "技术日报数据集_技术日报";
            reportDataSource3.Value = this.DDYLBindingSource;
            reportDataSource4.Name = "LTZN_质量日报_质量日报内容项";
            reportDataSource4.Value = this.DDYLBindingSource;
            reportDataSource5.Name = "LTZN_质量日报_机烧粒度内容项";
            reportDataSource5.Value = this.dDJSBindingSource;
            reportDataSource6.Name = "LTZN_质量日报_燃料指标内容项";
            reportDataSource6.Value = this.燃料指标内容项BindingSource;
            reportDataSource7.Name = "DDJS_DDJS";
            reportDataSource7.Value = this.dDJSBindingSource;
            reportDataSource8.Name = "LTZN_质量日报_套筒窑内容项";
            reportDataSource8.Value = this.dDTTYBindingSource;
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource1);
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource2);
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource3);
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource4);
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource5);
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource6);
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource7);
            this.reportViewer1.LocalReport.DataSources.Add(reportDataSource8);
            this.reportViewer1.LocalReport.ReportEmbeddedResource = "LTZN.质量日报.质量日报.rdlc";
            this.reportViewer1.Location = new System.Drawing.Point(0, 26);
            this.reportViewer1.Name = "reportViewer1";
            this.reportViewer1.Size = new System.Drawing.Size(823, 457);
            this.reportViewer1.TabIndex = 7;
            // 
            // webBrowser1
            // 
            this.webBrowser1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.webBrowser1.Location = new System.Drawing.Point(0, 26);
            this.webBrowser1.MinimumSize = new System.Drawing.Size(20, 20);
            this.webBrowser1.Name = "webBrowser1";
            this.webBrowser1.Size = new System.Drawing.Size(823, 457);
            this.webBrowser1.TabIndex = 6;
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripLabel2,
            this.riqiControl,
            this.toolStripSeparator4,
            this.btnQuery,
            this.toolStripSeparator6,
            this.btnSave,
            this.toolStripSeparator2,
            this.btnModel,
            this.toolStripSeparator1,
            this.btnPrint,
            this.toolStripSeparator3,
            this.btnExcel});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(823, 26);
            this.toolStrip1.TabIndex = 5;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripLabel2
            // 
            this.toolStripLabel2.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.toolStripLabel2.Name = "toolStripLabel2";
            this.toolStripLabel2.Size = new System.Drawing.Size(44, 23);
            this.toolStripLabel2.Text = "日期：";
            // 
            // riqiControl
            // 
            this.riqiControl.ControlWidth = 120;
            this.riqiControl.CustomFormat = null;
            this.riqiControl.Format = System.Windows.Forms.DateTimePickerFormat.Long;
            this.riqiControl.Name = "riqiControl";
            this.riqiControl.Size = new System.Drawing.Size(120, 23);
            this.riqiControl.Text = "2006年5月24日";
            this.riqiControl.Value = new System.DateTime(2006, 5, 24, 16, 23, 36, 1);
            // 
            // toolStripSeparator4
            // 
            this.toolStripSeparator4.Name = "toolStripSeparator4";
            this.toolStripSeparator4.Size = new System.Drawing.Size(6, 26);
            // 
            // btnQuery
            // 
            this.btnQuery.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btnQuery.Image = ((System.Drawing.Image)(resources.GetObject("btnQuery.Image")));
            this.btnQuery.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnQuery.Name = "btnQuery";
            this.btnQuery.Size = new System.Drawing.Size(60, 23);
            this.btnQuery.Text = "生成日报";
            this.btnQuery.Click += new System.EventHandler(this.btnQuery_Click);
            // 
            // toolStripSeparator6
            // 
            this.toolStripSeparator6.Name = "toolStripSeparator6";
            this.toolStripSeparator6.Size = new System.Drawing.Size(6, 26);
            // 
            // btnSave
            // 
            this.btnSave.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btnSave.Image = ((System.Drawing.Image)(resources.GetObject("btnSave.Image")));
            this.btnSave.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(36, 23);
            this.btnSave.Text = "保存";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 26);
            // 
            // btnModel
            // 
            this.btnModel.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btnModel.Image = ((System.Drawing.Image)(resources.GetObject("btnModel.Image")));
            this.btnModel.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnModel.Name = "btnModel";
            this.btnModel.Size = new System.Drawing.Size(60, 23);
            this.btnModel.Text = "报表模型";
            this.btnModel.Click += new System.EventHandler(this.btnModel_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 26);
            // 
            // btnPrint
            // 
            this.btnPrint.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btnPrint.Image = ((System.Drawing.Image)(resources.GetObject("btnPrint.Image")));
            this.btnPrint.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(36, 23);
            this.btnPrint.Text = "打印";
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(6, 26);
            // 
            // btnExcel
            // 
            this.btnExcel.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text;
            this.btnExcel.Image = ((System.Drawing.Image)(resources.GetObject("btnExcel.Image")));
            this.btnExcel.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(65, 23);
            this.btnExcel.Text = "导出Excel";
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // DDYLTableAdapter
            // 
            this.DDYLTableAdapter.ClearBeforeFill = true;
            // 
            // dDJSTableAdapter
            // 
            this.dDJSTableAdapter.ClearBeforeFill = true;
            // 
            // 焦炭指标内容项BindingSource
            // 
            this.焦炭指标内容项BindingSource.DataSource = typeof(LTZN.质量日报.焦炭指标内容项);
            // 
            // dDTTY
            // 
            this.dDTTY.DataSetName = "DDTTY";
            this.dDTTY.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // dDTTYBindingSource
            // 
            this.dDTTYBindingSource.DataMember = "DDTTY";
            this.dDTTYBindingSource.DataSource = this.dDTTY;
            // 
            // dDTTYTableAdapter
            // 
            this.dDTTYTableAdapter.ClearBeforeFill = true;
            // 
            // 质量日报Form
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(823, 483);
            this.Controls.Add(this.reportViewer1);
            this.Controls.Add(this.webBrowser1);
            this.Controls.Add(this.toolStrip1);
            this.Name = "质量日报Form";
            this.Text = "质量日报Form";
            this.Load += new System.EventHandler(this.质量日报Form_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DDYLBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dDJSBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dDJS)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.燃料指标内容项BindingSource)).EndInit();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.焦炭指标内容项BindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dDTTY)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dDTTYBindingSource)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private Microsoft.Reporting.WinForms.ReportViewer reportViewer1;
        private System.Windows.Forms.WebBrowser webBrowser1;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripLabel toolStripLabel2;
        private LTZN.Common.ToolStripDateTimePicker riqiControl;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator4;
        private System.Windows.Forms.ToolStripButton btnQuery;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator6;
        private System.Windows.Forms.ToolStripButton btnSave;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripButton btnModel;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton btnPrint;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
        private System.Windows.Forms.ToolStripButton btnExcel;
        private System.Windows.Forms.BindingSource DDYLBindingSource;
        private LTZN.质量日报.质量日报数据集TableAdapters.DDYLTableAdapter DDYLTableAdapter;
        private System.Windows.Forms.BindingSource dDJSBindingSource;
        private DDJS dDJS;
        private LTZN.质量日报.DDJSTableAdapters.DDJSTableAdapter dDJSTableAdapter;
        private System.Windows.Forms.BindingSource 焦炭指标内容项BindingSource;
        private System.Windows.Forms.BindingSource 燃料指标内容项BindingSource;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private DDTTY dDTTY;
        private System.Windows.Forms.BindingSource dDTTYBindingSource;
        private LTZN.质量日报.DDTTYTableAdapters.DDTTYTableAdapter dDTTYTableAdapter;
    }
}