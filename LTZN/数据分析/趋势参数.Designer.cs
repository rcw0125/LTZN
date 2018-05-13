namespace LTZN.数据分析
{
    partial class 趋势参数
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
            this.cANSHUDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cANSHUMAXDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cANSHUMINDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.yMAXDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.yMINDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lTQSPARABindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.dataSet1 = new LTZN.DataSet1();
            this.lT_QSPARATableAdapter = new LTZN.DataSet1TableAdapters.LT_QSPARATableAdapter();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lTQSPARABindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.Control;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.gAOLUDataGridViewTextBoxColumn,
            this.cANSHUDataGridViewTextBoxColumn,
            this.cANSHUMAXDataGridViewTextBoxColumn,
            this.cANSHUMINDataGridViewTextBoxColumn,
            this.yMAXDataGridViewTextBoxColumn,
            this.yMINDataGridViewTextBoxColumn});
            this.dataGridView1.DataSource = this.lTQSPARABindingSource;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(737, 495);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellEndEdit);
            // 
            // gAOLUDataGridViewTextBoxColumn
            // 
            this.gAOLUDataGridViewTextBoxColumn.DataPropertyName = "GAOLU";
            this.gAOLUDataGridViewTextBoxColumn.HeaderText = "高炉";
            this.gAOLUDataGridViewTextBoxColumn.Name = "gAOLUDataGridViewTextBoxColumn";
            // 
            // cANSHUDataGridViewTextBoxColumn
            // 
            this.cANSHUDataGridViewTextBoxColumn.DataPropertyName = "CANSHU";
            this.cANSHUDataGridViewTextBoxColumn.HeaderText = "参数名称";
            this.cANSHUDataGridViewTextBoxColumn.Name = "cANSHUDataGridViewTextBoxColumn";
            // 
            // cANSHUMAXDataGridViewTextBoxColumn
            // 
            this.cANSHUMAXDataGridViewTextBoxColumn.DataPropertyName = "CANSHUMAX";
            this.cANSHUMAXDataGridViewTextBoxColumn.HeaderText = "参数上限";
            this.cANSHUMAXDataGridViewTextBoxColumn.Name = "cANSHUMAXDataGridViewTextBoxColumn";
            // 
            // cANSHUMINDataGridViewTextBoxColumn
            // 
            this.cANSHUMINDataGridViewTextBoxColumn.DataPropertyName = "CANSHUMIN";
            this.cANSHUMINDataGridViewTextBoxColumn.HeaderText = "参数下限";
            this.cANSHUMINDataGridViewTextBoxColumn.Name = "cANSHUMINDataGridViewTextBoxColumn";
            // 
            // yMAXDataGridViewTextBoxColumn
            // 
            this.yMAXDataGridViewTextBoxColumn.DataPropertyName = "YMAX";
            this.yMAXDataGridViewTextBoxColumn.HeaderText = "Y轴最大值";
            this.yMAXDataGridViewTextBoxColumn.Name = "yMAXDataGridViewTextBoxColumn";
            // 
            // yMINDataGridViewTextBoxColumn
            // 
            this.yMINDataGridViewTextBoxColumn.DataPropertyName = "YMIN";
            this.yMINDataGridViewTextBoxColumn.HeaderText = "Y轴最小值";
            this.yMINDataGridViewTextBoxColumn.Name = "yMINDataGridViewTextBoxColumn";
            // 
            // lTQSPARABindingSource
            // 
            this.lTQSPARABindingSource.DataMember = "LT_QSPARA";
            this.lTQSPARABindingSource.DataSource = this.dataSet1;
            // 
            // dataSet1
            // 
            this.dataSet1.DataSetName = "DataSet1";
            this.dataSet1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // lT_QSPARATableAdapter
            // 
            this.lT_QSPARATableAdapter.ClearBeforeFill = true;
            // 
            // 趋势参数
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(737, 495);
            this.Controls.Add(this.dataGridView1);
            this.Name = "趋势参数";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "趋势参数";
            this.Load += new System.EventHandler(this.趋势参数_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lTQSPARABindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private DataSet1 dataSet1;
        private System.Windows.Forms.BindingSource lTQSPARABindingSource;
        private LTZN.DataSet1TableAdapters.LT_QSPARATableAdapter lT_QSPARATableAdapter;
        private System.Windows.Forms.DataGridViewTextBoxColumn gAOLUDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn cANSHUDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn cANSHUMAXDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn cANSHUMINDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn yMAXDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn yMINDataGridViewTextBoxColumn;
    }
}