namespace LTZN
{
    partial class 出铁方案
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
            this.fANAMEDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.xUHAODataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.sJDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cHUTIEFANGANBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.ckfa = new LTZN.ckfa();
            this.保存 = new System.Windows.Forms.Button();
            this.cHUTIEFANGANTableAdapter = new LTZN.ckfaTableAdapters.CHUTIEFANGANTableAdapter();
            this.ckfa1 = new LTZN.ckfa();
            this.ckfa2 = new LTZN.ckfa();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.cHUTIEFANGANBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ckfa)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ckfa1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.ckfa2)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.fANAMEDataGridViewTextBoxColumn,
            this.xUHAODataGridViewTextBoxColumn,
            this.sJDataGridViewTextBoxColumn});
            this.dataGridView1.DataSource = this.cHUTIEFANGANBindingSource;
            this.dataGridView1.Location = new System.Drawing.Point(0, 1);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(343, 471);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.DataError += new System.Windows.Forms.DataGridViewDataErrorEventHandler(this.dataGridView1_DataError_1);
            // 
            // fANAMEDataGridViewTextBoxColumn
            // 
            this.fANAMEDataGridViewTextBoxColumn.DataPropertyName = "FANAME";
            this.fANAMEDataGridViewTextBoxColumn.HeaderText = "方案名称";
            this.fANAMEDataGridViewTextBoxColumn.Name = "fANAMEDataGridViewTextBoxColumn";
            // 
            // xUHAODataGridViewTextBoxColumn
            // 
            this.xUHAODataGridViewTextBoxColumn.DataPropertyName = "XUHAO";
            this.xUHAODataGridViewTextBoxColumn.HeaderText = "序号";
            this.xUHAODataGridViewTextBoxColumn.Name = "xUHAODataGridViewTextBoxColumn";
            // 
            // sJDataGridViewTextBoxColumn
            // 
            this.sJDataGridViewTextBoxColumn.DataPropertyName = "SJ";
            this.sJDataGridViewTextBoxColumn.HeaderText = "时间";
            this.sJDataGridViewTextBoxColumn.Name = "sJDataGridViewTextBoxColumn";
            // 
            // cHUTIEFANGANBindingSource
            // 
            this.cHUTIEFANGANBindingSource.DataMember = "CHUTIEFANGAN";
            this.cHUTIEFANGANBindingSource.DataSource = this.ckfa;
            // 
            // ckfa
            // 
            this.ckfa.DataSetName = "ckfa";
            this.ckfa.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // 保存
            // 
            this.保存.Location = new System.Drawing.Point(268, 478);
            this.保存.Name = "保存";
            this.保存.Size = new System.Drawing.Size(75, 28);
            this.保存.TabIndex = 1;
            this.保存.Text = "保存";
            this.保存.UseVisualStyleBackColor = true;
            this.保存.Click += new System.EventHandler(this.保存_Click);
            // 
            // cHUTIEFANGANTableAdapter
            // 
            this.cHUTIEFANGANTableAdapter.ClearBeforeFill = true;
            // 
            // ckfa1
            // 
            this.ckfa1.DataSetName = "ckfa";
            this.ckfa1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // ckfa2
            // 
            this.ckfa2.DataSetName = "ckfa";
            this.ckfa2.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // 出铁方案
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(344, 509);
            this.Controls.Add(this.保存);
            this.Controls.Add(this.dataGridView1);
            this.Name = "出铁方案";
            this.Text = "出铁方案";
            this.Load += new System.EventHandler(this.出铁方案_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.cHUTIEFANGANBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ckfa)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ckfa1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.ckfa2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button 保存;
        private ckfa ckfa;
        private System.Windows.Forms.BindingSource cHUTIEFANGANBindingSource;
        private LTZN.ckfaTableAdapters.CHUTIEFANGANTableAdapter cHUTIEFANGANTableAdapter;
        private System.Windows.Forms.DataGridViewTextBoxColumn fANAMEDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn xUHAODataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn sJDataGridViewTextBoxColumn;
        private ckfa ckfa1;
        private ckfa ckfa2;
    }
}