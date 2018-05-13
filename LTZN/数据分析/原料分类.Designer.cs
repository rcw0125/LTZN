namespace LTZN.数据分析
{
    partial class 原料分类
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
            this.yLZLDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.yLMCDataGridViewTextBoxColumn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.yLFLBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.dataSet2 = new LTZN.DataSet2();
            this.yLFLTableAdapter = new LTZN.DataSet2TableAdapters.YLFLTableAdapter();
            this.button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.yLFLBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet2)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.yLZLDataGridViewTextBoxColumn,
            this.yLMCDataGridViewTextBoxColumn});
            this.dataGridView1.DataSource = this.yLFLBindingSource;
            this.dataGridView1.Location = new System.Drawing.Point(12, 12);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 23;
            this.dataGridView1.Size = new System.Drawing.Size(766, 491);
            this.dataGridView1.TabIndex = 0;
            // 
            // yLZLDataGridViewTextBoxColumn
            // 
            this.yLZLDataGridViewTextBoxColumn.DataPropertyName = "YLZL";
            this.yLZLDataGridViewTextBoxColumn.HeaderText = "原料种类";
            this.yLZLDataGridViewTextBoxColumn.Name = "yLZLDataGridViewTextBoxColumn";
            // 
            // yLMCDataGridViewTextBoxColumn
            // 
            this.yLMCDataGridViewTextBoxColumn.DataPropertyName = "YLMC";
            this.yLMCDataGridViewTextBoxColumn.HeaderText = "原料名称";
            this.yLMCDataGridViewTextBoxColumn.Name = "yLMCDataGridViewTextBoxColumn";
            this.yLMCDataGridViewTextBoxColumn.Width = 200;
            // 
            // yLFLBindingSource
            // 
            this.yLFLBindingSource.DataMember = "YLFL";
            this.yLFLBindingSource.DataSource = this.dataSet2;
            // 
            // dataSet2
            // 
            this.dataSet2.DataSetName = "DataSet2";
            this.dataSet2.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // yLFLTableAdapter
            // 
            this.yLFLTableAdapter.ClearBeforeFill = true;
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button1.Location = new System.Drawing.Point(629, 519);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(92, 28);
            this.button1.TabIndex = 1;
            this.button1.Text = "保存";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // 原料分类
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(820, 559);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dataGridView1);
            this.Name = "原料分类";
            this.Text = "原料分类";
            this.Load += new System.EventHandler(this.原料分类_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.yLFLBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private DataSet2 dataSet2;
        private System.Windows.Forms.BindingSource yLFLBindingSource;
        private LTZN.DataSet2TableAdapters.YLFLTableAdapter yLFLTableAdapter;
        private System.Windows.Forms.DataGridViewTextBoxColumn yLZLDataGridViewTextBoxColumn;
        private System.Windows.Forms.DataGridViewTextBoxColumn yLMCDataGridViewTextBoxColumn;
        private System.Windows.Forms.Button button1;
    }
}