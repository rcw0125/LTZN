namespace LTZN
{
    partial class 出铁计划
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
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.gAOLUDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.kSRQDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.fANGANDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cHUTIEJIHUABindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.ctjh = new LTZN.ctjh();
            this.cHUTIEJIHUATableAdapter = new LTZN.ctjhTableAdapters.CHUTIEJIHUATableAdapter();
            this.button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.cHUTIEJIHUABindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ctjh)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.gAOLUDataGridViewTextBoxColumn,
            this.kSRQDataGridViewTextBoxColumn,
            this.fANGANDataGridViewTextBoxColumn});
            this.dataGridView1.DataSource = this.cHUTIEJIHUABindingSource;
            this.dataGridView1.Location = new System.Drawing.Point(0, -1);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(343, 509);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.DataError += new System.Windows.Forms.DataGridViewDataErrorEventHandler(this.dataGridView1_DataError);
            // 
            // gAOLUDataGridViewTextBoxColumn
            // 
            this.gAOLUDataGridViewTextBoxColumn.DataPropertyName = "GAOLU";
            this.gAOLUDataGridViewTextBoxColumn.HeaderText = "高炉";
            this.gAOLUDataGridViewTextBoxColumn.Name = "gAOLUDataGridViewTextBoxColumn";
            // 
            // kSRQDataGridViewTextBoxColumn
            // 
            this.kSRQDataGridViewTextBoxColumn.DataPropertyName = "KSRQ";
            this.kSRQDataGridViewTextBoxColumn.HeaderText = "开始时间";
            this.kSRQDataGridViewTextBoxColumn.Name = "kSRQDataGridViewTextBoxColumn";
            // 
            // fANGANDataGridViewTextBoxColumn
            // 
            this.fANGANDataGridViewTextBoxColumn.DataPropertyName = "FANGAN";
            this.fANGANDataGridViewTextBoxColumn.HeaderText = "方案名称";
            this.fANGANDataGridViewTextBoxColumn.Name = "fANGANDataGridViewTextBoxColumn";
            // 
            // cHUTIEJIHUABindingSource
            // 
            this.cHUTIEJIHUABindingSource.DataMember = "CHUTIEJIHUA";
            this.cHUTIEJIHUABindingSource.DataSource = this.ctjh;
            // 
            // ctjh
            // 
            this.ctjh.DataSetName = "ctjh";
            this.ctjh.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // cHUTIEJIHUATableAdapter
            // 
            this.cHUTIEJIHUATableAdapter.ClearBeforeFill = true;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(267, 514);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 30);
            this.button1.TabIndex = 1;
            this.button1.Text = "保存";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // 出铁计划
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(345, 542);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dataGridView1);
            this.Name = "出铁计划";
            this.Text = "出铁计划";
            this.Load += new System.EventHandler(this.出铁计划_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.cHUTIEJIHUABindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ctjh)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private ctjh ctjh;
        private System.Windows.Forms.BindingSource cHUTIEJIHUABindingSource;
        private LTZN.ctjhTableAdapters.CHUTIEJIHUATableAdapter cHUTIEJIHUATableAdapter;
        private System.Windows.Forms.DataGridViewTextBoxColumn gAOLUDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn kSRQDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn fANGANDataGridViewTextBoxColumn;
        private System.Windows.Forms.Button button1;
    }
}