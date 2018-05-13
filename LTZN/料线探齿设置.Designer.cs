namespace LTZN
{
    partial class 料线探齿设置
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(料线探齿设置));
            this.c1FlexGrid_xiaohao = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.lxbz = new LTZN.lxbz();
            this.lXBIAOZHIBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.lXBIAOZHITableAdapter = new LTZN.lxbzTableAdapters.LXBIAOZHITableAdapter();
            this.tableAdapterManager = new LTZN.lxbzTableAdapters.TableAdapterManager();
            this.保存 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.c1FlexGrid_xiaohao)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lxbz)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.lXBIAOZHIBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // c1FlexGrid_xiaohao
            // 
            this.c1FlexGrid_xiaohao.AllowDragging = C1.Win.C1FlexGrid.AllowDraggingEnum.None;
            this.c1FlexGrid_xiaohao.AllowSorting = C1.Win.C1FlexGrid.AllowSortingEnum.None;
            this.c1FlexGrid_xiaohao.AutoClipboard = true;
            this.c1FlexGrid_xiaohao.AutoGenerateColumns = false;
            this.c1FlexGrid_xiaohao.ColumnInfo = resources.GetString("c1FlexGrid_xiaohao.ColumnInfo");
            this.c1FlexGrid_xiaohao.DataSource = this.lXBIAOZHIBindingSource;
            this.c1FlexGrid_xiaohao.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.c1FlexGrid_xiaohao.KeyActionEnter = C1.Win.C1FlexGrid.KeyActionEnum.MoveAcross;
            this.c1FlexGrid_xiaohao.Location = new System.Drawing.Point(0, 0);
            this.c1FlexGrid_xiaohao.Name = "c1FlexGrid_xiaohao";
            this.c1FlexGrid_xiaohao.Rows.Count = 1;
            this.c1FlexGrid_xiaohao.Rows.DefaultSize = 23;
            this.c1FlexGrid_xiaohao.SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.Row;
            this.c1FlexGrid_xiaohao.Size = new System.Drawing.Size(455, 198);
            this.c1FlexGrid_xiaohao.StyleInfo = resources.GetString("c1FlexGrid_xiaohao.StyleInfo");
            this.c1FlexGrid_xiaohao.TabIndex = 14;
            this.c1FlexGrid_xiaohao.Tag = "";
            // 
            // lxbz
            // 
            this.lxbz.DataSetName = "lxbz";
            this.lxbz.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // lXBIAOZHIBindingSource
            // 
            this.lXBIAOZHIBindingSource.DataMember = "LXBIAOZHI";
            this.lXBIAOZHIBindingSource.DataSource = this.lxbz;
            // 
            // lXBIAOZHITableAdapter
            // 
            this.lXBIAOZHITableAdapter.ClearBeforeFill = true;
            // 
            // tableAdapterManager
            // 
            this.tableAdapterManager.BackupDataSetBeforeUpdate = false;
            this.tableAdapterManager.LXBIAOZHITableAdapter = this.lXBIAOZHITableAdapter;
            this.tableAdapterManager.UpdateOrder = LTZN.lxbzTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete;
            // 
            // 保存
            // 
            this.保存.Location = new System.Drawing.Point(346, 200);
            this.保存.Name = "保存";
            this.保存.Size = new System.Drawing.Size(75, 23);
            this.保存.TabIndex = 15;
            this.保存.Text = "保存";
            this.保存.UseVisualStyleBackColor = true;
            this.保存.Click += new System.EventHandler(this.保存_Click);
            // 
            // 料线探齿设置
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(452, 235);
            this.Controls.Add(this.保存);
            this.Controls.Add(this.c1FlexGrid_xiaohao);
            this.Name = "料线探齿设置";
            this.Text = "料线探齿设置";
            this.Load += new System.EventHandler(this.料线探齿设置_Load);
            ((System.ComponentModel.ISupportInitialize)(this.c1FlexGrid_xiaohao)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lxbz)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.lXBIAOZHIBindingSource)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private C1.Win.C1FlexGrid.C1FlexGrid c1FlexGrid_xiaohao;
        private lxbz lxbz;
        private System.Windows.Forms.BindingSource lXBIAOZHIBindingSource;
        private LTZN.lxbzTableAdapters.LXBIAOZHITableAdapter lXBIAOZHITableAdapter;
        private LTZN.lxbzTableAdapters.TableAdapterManager tableAdapterManager;
        private System.Windows.Forms.Button 保存;

    }
}