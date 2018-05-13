namespace LTZN.质量日报
{
    partial class Form1
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
            System.Windows.Forms.DataGridView dataGridView1;
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.label1 = new System.Windows.Forms.Label();
            this.dDYLBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.dataSet3 = new LTZN.DataSet3();
            this.dDYLTableAdapter = new LTZN.DataSet3TableAdapters.DDYLTableAdapter();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.文件 = new System.Windows.Forms.ToolStripLabel();
            this.导出Excel = new System.Windows.Forms.ToolStripLabel();
            this.名称 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.时间 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TFE = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.FeO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SiO2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.CaO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MgO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Al2O3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.S = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.P = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TiO2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.R = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Cr = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Ni = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.As = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MnO = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Pb = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Zn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.碱金属 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.V2O5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Cu = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Mo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Sn = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Sb = new System.Windows.Forms.DataGridViewTextBoxColumn();
            dataGridView1 = new System.Windows.Forms.DataGridView();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dDYLBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet3)).BeginInit();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.toolStrip1);
            this.splitContainer1.Panel1.Controls.Add(this.label1);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(dataGridView1);
            this.splitContainer1.Size = new System.Drawing.Size(1481, 654);
            this.splitContainer1.SplitterDistance = 99;
            this.splitContainer1.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(650, 36);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(143, 33);
            this.label1.TabIndex = 0;
            this.label1.Text = "质量日报";
            // 
            // dataGridView1
            // 
            dataGridView1.AutoGenerateColumns = false;
            dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.名称,
            this.时间,
            this.TFE,
            this.FeO,
            this.SiO2,
            this.CaO,
            this.MgO,
            this.Al2O3,
            this.S,
            this.P,
            this.TiO2,
            this.R,
            this.Cr,
            this.Ni,
            this.As,
            this.MnO,
            this.Pb,
            this.Zn,
            this.碱金属,
            this.V2O5,
            this.Cu,
            this.Mo,
            this.Sn,
            this.Sb});
            dataGridView1.DataSource = this.dDYLBindingSource;
            dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            dataGridView1.Location = new System.Drawing.Point(0, 0);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.RowTemplate.Height = 23;
            dataGridView1.Size = new System.Drawing.Size(1481, 551);
            dataGridView1.TabIndex = 0;
            // 
            // dDYLBindingSource
            // 
            this.dDYLBindingSource.DataMember = "DDYL";
            this.dDYLBindingSource.DataSource = this.dataSet3;
            // 
            // dataSet3
            // 
            this.dataSet3.DataSetName = "DataSet3";
            this.dataSet3.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // dDYLTableAdapter
            // 
            this.dDYLTableAdapter.ClearBeforeFill = true;
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.文件,
            this.导出Excel});
            this.toolStrip1.Location = new System.Drawing.Point(0, 0);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(1481, 25);
            this.toolStrip1.TabIndex = 1;
            this.toolStrip1.Text = "文件";
            // 
            // 文件
            // 
            this.文件.Name = "文件";
            this.文件.Size = new System.Drawing.Size(56, 22);
            this.文件.Text = "生成日报";
            // 
            // 导出Excel
            // 
            this.导出Excel.Name = "导出Excel";
            this.导出Excel.Size = new System.Drawing.Size(61, 22);
            this.导出Excel.Text = "导出Excel";
            // 
            // 名称
            // 
            this.名称.HeaderText = "名称";
            this.名称.Name = "名称";
            // 
            // 时间
            // 
            this.时间.HeaderText = "时间";
            this.时间.Name = "时间";
            // 
            // TFE
            // 
            this.TFE.HeaderText = "TFE";
            this.TFE.Name = "TFE";
            // 
            // FeO
            // 
            this.FeO.HeaderText = "FeO";
            this.FeO.Name = "FeO";
            // 
            // SiO2
            // 
            this.SiO2.HeaderText = "SiO2";
            this.SiO2.Name = "SiO2";
            // 
            // CaO
            // 
            this.CaO.HeaderText = "CaO";
            this.CaO.Name = "CaO";
            // 
            // MgO
            // 
            this.MgO.HeaderText = "MgO";
            this.MgO.Name = "MgO";
            // 
            // Al2O3
            // 
            this.Al2O3.HeaderText = "Al2O3";
            this.Al2O3.Name = "Al2O3";
            // 
            // S
            // 
            this.S.HeaderText = "S";
            this.S.Name = "S";
            // 
            // P
            // 
            this.P.HeaderText = "P";
            this.P.Name = "P";
            // 
            // TiO2
            // 
            this.TiO2.HeaderText = "TiO2";
            this.TiO2.Name = "TiO2";
            // 
            // R
            // 
            this.R.HeaderText = "R";
            this.R.Name = "R";
            // 
            // Cr
            // 
            this.Cr.HeaderText = "Cr";
            this.Cr.Name = "Cr";
            // 
            // Ni
            // 
            this.Ni.HeaderText = "Ni";
            this.Ni.Name = "Ni";
            // 
            // As
            // 
            this.As.HeaderText = "As";
            this.As.Name = "As";
            // 
            // MnO
            // 
            this.MnO.HeaderText = "MnO";
            this.MnO.Name = "MnO";
            // 
            // Pb
            // 
            this.Pb.HeaderText = "Pb";
            this.Pb.Name = "Pb";
            // 
            // Zn
            // 
            this.Zn.HeaderText = "Zn";
            this.Zn.Name = "Zn";
            // 
            // 碱金属
            // 
            this.碱金属.HeaderText = "碱金属";
            this.碱金属.Name = "碱金属";
            // 
            // V2O5
            // 
            this.V2O5.HeaderText = "V2O5";
            this.V2O5.Name = "V2O5";
            // 
            // Cu
            // 
            this.Cu.HeaderText = "Cu";
            this.Cu.Name = "Cu";
            // 
            // Mo
            // 
            this.Mo.HeaderText = "Mo";
            this.Mo.Name = "Mo";
            // 
            // Sn
            // 
            this.Sn.HeaderText = "Sn";
            this.Sn.Name = "Sn";
            // 
            // Sb
            // 
            this.Sb.HeaderText = "Sb";
            this.Sb.Name = "Sb";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1481, 654);
            this.Controls.Add(this.splitContainer1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel1.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dDYLBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet3)).EndInit();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.Label label1;
        private DataSet3 dataSet3;
        private System.Windows.Forms.BindingSource dDYLBindingSource;
        private LTZN.DataSet3TableAdapters.DDYLTableAdapter dDYLTableAdapter;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripLabel 文件;
        private System.Windows.Forms.ToolStripLabel 导出Excel;
        private System.Windows.Forms.DataGridViewTextBoxColumn 名称;
        private System.Windows.Forms.DataGridViewTextBoxColumn 时间;
        private System.Windows.Forms.DataGridViewTextBoxColumn TFE;
        private System.Windows.Forms.DataGridViewTextBoxColumn FeO;
        private System.Windows.Forms.DataGridViewTextBoxColumn SiO2;
        private System.Windows.Forms.DataGridViewTextBoxColumn CaO;
        private System.Windows.Forms.DataGridViewTextBoxColumn MgO;
        private System.Windows.Forms.DataGridViewTextBoxColumn Al2O3;
        private System.Windows.Forms.DataGridViewTextBoxColumn S;
        private System.Windows.Forms.DataGridViewTextBoxColumn P;
        private System.Windows.Forms.DataGridViewTextBoxColumn TiO2;
        private System.Windows.Forms.DataGridViewTextBoxColumn R;
        private System.Windows.Forms.DataGridViewTextBoxColumn Cr;
        private System.Windows.Forms.DataGridViewTextBoxColumn Ni;
        private System.Windows.Forms.DataGridViewTextBoxColumn As;
        private System.Windows.Forms.DataGridViewTextBoxColumn MnO;
        private System.Windows.Forms.DataGridViewTextBoxColumn Pb;
        private System.Windows.Forms.DataGridViewTextBoxColumn Zn;
        private System.Windows.Forms.DataGridViewTextBoxColumn 碱金属;
        private System.Windows.Forms.DataGridViewTextBoxColumn V2O5;
        private System.Windows.Forms.DataGridViewTextBoxColumn Cu;
        private System.Windows.Forms.DataGridViewTextBoxColumn Mo;
        private System.Windows.Forms.DataGridViewTextBoxColumn Sn;
        private System.Windows.Forms.DataGridViewTextBoxColumn Sb;
    }
}