namespace LTZN
{
    partial class CalModelForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CalModelForm));
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.treeView1 = new System.Windows.Forms.TreeView();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.btnAddModel = new System.Windows.Forms.ToolStripMenuItem();
            this.btnAddOrg = new System.Windows.Forms.ToolStripMenuItem();
            this.btnAddTag = new System.Windows.Forms.ToolStripMenuItem();
            this.btnSave = new System.Windows.Forms.ToolStripMenuItem();
            this.btnDeleteModel = new System.Windows.Forms.ToolStripMenuItem();
            this.btnDeleteOrg = new System.Windows.Forms.ToolStripMenuItem();
            this.btnCalc = new System.Windows.Forms.ToolStripMenuItem();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.c1FlexGrid1 = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.contextMenuStrip2 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.btnAddTag2 = new System.Windows.Forms.ToolStripMenuItem();
            this.btnDeleteTag = new System.Windows.Forms.ToolStripMenuItem();
            this.btnCopy = new System.Windows.Forms.ToolStripMenuItem();
            this.btnPaste = new System.Windows.Forms.ToolStripMenuItem();
            this.btnSumForma = new System.Windows.Forms.ToolStripMenuItem();
            this.btnFenTanForma = new System.Windows.Forms.ToolStripMenuItem();
            this.设定累计公式ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.上移ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.下移ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.flowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.ckbViewType = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cmbFilterTyp = new System.Windows.Forms.ComboBox();
            this.txtFilter = new System.Windows.Forms.TextBox();
            this.buttonOk = new System.Windows.Forms.Button();
            this.buttonCal = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.contextMenuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1FlexGrid1)).BeginInit();
            this.contextMenuStrip2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.flowLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.treeView1);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.c1FlexGrid1);
            this.splitContainer1.Panel2.Controls.Add(this.flowLayoutPanel1);
            this.splitContainer1.Size = new System.Drawing.Size(1094, 445);
            this.splitContainer1.SplitterDistance = 204;
            this.splitContainer1.TabIndex = 0;
            // 
            // treeView1
            // 
            this.treeView1.AllowDrop = true;
            this.treeView1.ContextMenuStrip = this.contextMenuStrip1;
            this.treeView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeView1.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.treeView1.ImageIndex = 5;
            this.treeView1.ImageList = this.imageList1;
            this.treeView1.ItemHeight = 20;
            this.treeView1.Location = new System.Drawing.Point(0, 0);
            this.treeView1.Name = "treeView1";
            this.treeView1.SelectedImageIndex = 0;
            this.treeView1.Size = new System.Drawing.Size(204, 445);
            this.treeView1.TabIndex = 3;
            this.treeView1.AfterLabelEdit += new System.Windows.Forms.NodeLabelEditEventHandler(this.treeView1_AfterLabelEdit);
            this.treeView1.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeView1_AfterSelect);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnAddModel,
            this.btnAddOrg,
            this.btnAddTag,
            this.btnSave,
            this.btnDeleteModel,
            this.btnDeleteOrg,
            this.btnCalc});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(143, 158);
            // 
            // btnAddModel
            // 
            this.btnAddModel.Name = "btnAddModel";
            this.btnAddModel.Size = new System.Drawing.Size(142, 22);
            this.btnAddModel.Text = "添加模型";
            this.btnAddModel.Click += new System.EventHandler(this.添加模型ToolStripMenuItem_Click);
            // 
            // btnAddOrg
            // 
            this.btnAddOrg.Name = "btnAddOrg";
            this.btnAddOrg.Size = new System.Drawing.Size(142, 22);
            this.btnAddOrg.Text = "添加组织";
            this.btnAddOrg.Click += new System.EventHandler(this.添加组织结构ToolStripMenuItem_Click);
            // 
            // btnAddTag
            // 
            this.btnAddTag.Name = "btnAddTag";
            this.btnAddTag.Size = new System.Drawing.Size(142, 22);
            this.btnAddTag.Text = "添加标签系列";
            this.btnAddTag.Click += new System.EventHandler(this.添加标签ToolStripMenuItem_Click);
            // 
            // btnSave
            // 
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(142, 22);
            this.btnSave.Text = "保存";
            this.btnSave.Click += new System.EventHandler(this.保存ToolStripMenuItem_Click);
            // 
            // btnDeleteModel
            // 
            this.btnDeleteModel.Name = "btnDeleteModel";
            this.btnDeleteModel.Size = new System.Drawing.Size(142, 22);
            this.btnDeleteModel.Text = "删除模型";
            this.btnDeleteModel.Click += new System.EventHandler(this.删除模型ToolStripMenuItem_Click);
            // 
            // btnDeleteOrg
            // 
            this.btnDeleteOrg.Name = "btnDeleteOrg";
            this.btnDeleteOrg.Size = new System.Drawing.Size(142, 22);
            this.btnDeleteOrg.Text = "删除组织";
            this.btnDeleteOrg.Click += new System.EventHandler(this.删除组织ToolStripMenuItem_Click);
            // 
            // btnCalc
            // 
            this.btnCalc.Name = "btnCalc";
            this.btnCalc.Size = new System.Drawing.Size(142, 22);
            this.btnCalc.Text = "计算";
            this.btnCalc.Click += new System.EventHandler(this.计算ToolStripMenuItem_Click);
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "access.ico");
            this.imageList1.Images.SetKeyName(1, "bar_default.png");
            this.imageList1.Images.SetKeyName(2, "bnexec.gif");
            this.imageList1.Images.SetKeyName(3, "bninexec.gif");
            this.imageList1.Images.SetKeyName(4, "product.png");
            this.imageList1.Images.SetKeyName(5, "usb.ico");
            this.imageList1.Images.SetKeyName(6, "superkaramba.ico");
            // 
            // c1FlexGrid1
            // 
            this.c1FlexGrid1.AllowDragging = C1.Win.C1FlexGrid.AllowDraggingEnum.Rows;
            this.c1FlexGrid1.AllowSorting = C1.Win.C1FlexGrid.AllowSortingEnum.None;
            this.c1FlexGrid1.AutoClipboard = true;
            this.c1FlexGrid1.AutoGenerateColumns = false;
            this.c1FlexGrid1.AutoResize = false;
            this.c1FlexGrid1.ColumnInfo = resources.GetString("c1FlexGrid1.ColumnInfo");
            this.c1FlexGrid1.ContextMenuStrip = this.contextMenuStrip2;
            this.c1FlexGrid1.DataSource = this.bindingSource1;
            this.c1FlexGrid1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.c1FlexGrid1.Font = new System.Drawing.Font("宋体", 12F);
            this.c1FlexGrid1.ImeMode = System.Windows.Forms.ImeMode.Alpha;
            this.c1FlexGrid1.KeyActionEnter = C1.Win.C1FlexGrid.KeyActionEnum.MoveAcross;
            this.c1FlexGrid1.Location = new System.Drawing.Point(0, 36);
            this.c1FlexGrid1.Name = "c1FlexGrid1";
            this.c1FlexGrid1.Rows.Count = 1;
            this.c1FlexGrid1.Rows.DefaultSize = 23;
            this.c1FlexGrid1.SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.RowRange;
            this.c1FlexGrid1.Size = new System.Drawing.Size(886, 409);
            this.c1FlexGrid1.StyleInfo = resources.GetString("c1FlexGrid1.StyleInfo");
            this.c1FlexGrid1.TabIndex = 14;
            this.c1FlexGrid1.Tag = "";
            // 
            // contextMenuStrip2
            // 
            this.contextMenuStrip2.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.btnAddTag2,
            this.btnDeleteTag,
            this.btnCopy,
            this.btnPaste,
            this.btnSumForma,
            this.btnFenTanForma,
            this.设定累计公式ToolStripMenuItem,
            this.上移ToolStripMenuItem,
            this.下移ToolStripMenuItem});
            this.contextMenuStrip2.Name = "contextMenuStrip2";
            this.contextMenuStrip2.Size = new System.Drawing.Size(154, 224);
            // 
            // btnAddTag2
            // 
            this.btnAddTag2.Name = "btnAddTag2";
            this.btnAddTag2.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Alt | System.Windows.Forms.Keys.A)));
            this.btnAddTag2.Size = new System.Drawing.Size(153, 22);
            this.btnAddTag2.Text = "添加标签";
            this.btnAddTag2.Click += new System.EventHandler(this.添加标签_Click);
            // 
            // btnDeleteTag
            // 
            this.btnDeleteTag.Name = "btnDeleteTag";
            this.btnDeleteTag.Size = new System.Drawing.Size(153, 22);
            this.btnDeleteTag.Text = "删除标签";
            this.btnDeleteTag.Click += new System.EventHandler(this.删除标签ToolStripMenuItem1_Click);
            // 
            // btnCopy
            // 
            this.btnCopy.Name = "btnCopy";
            this.btnCopy.Size = new System.Drawing.Size(153, 22);
            this.btnCopy.Text = "复制";
            this.btnCopy.Click += new System.EventHandler(this.复制ToolStripMenuItem_Click);
            // 
            // btnPaste
            // 
            this.btnPaste.Name = "btnPaste";
            this.btnPaste.Size = new System.Drawing.Size(153, 22);
            this.btnPaste.Text = "粘贴";
            this.btnPaste.Click += new System.EventHandler(this.粘贴ToolStripMenuItem_Click);
            // 
            // btnSumForma
            // 
            this.btnSumForma.Name = "btnSumForma";
            this.btnSumForma.Size = new System.Drawing.Size(153, 22);
            this.btnSumForma.Text = "设定求和公式";
            this.btnSumForma.Click += new System.EventHandler(this.设定求和公式ToolStripMenuItem_Click);
            // 
            // btnFenTanForma
            // 
            this.btnFenTanForma.Name = "btnFenTanForma";
            this.btnFenTanForma.Size = new System.Drawing.Size(153, 22);
            this.btnFenTanForma.Text = "设定分摊公式";
            this.btnFenTanForma.Click += new System.EventHandler(this.设定分摊公式ToolStripMenuItem_Click);
            // 
            // 设定累计公式ToolStripMenuItem
            // 
            this.设定累计公式ToolStripMenuItem.Name = "设定累计公式ToolStripMenuItem";
            this.设定累计公式ToolStripMenuItem.Size = new System.Drawing.Size(153, 22);
            this.设定累计公式ToolStripMenuItem.Text = "设定累计公式";
            this.设定累计公式ToolStripMenuItem.Click += new System.EventHandler(this.设定累计公式ToolStripMenuItem_Click);
            // 
            // 上移ToolStripMenuItem
            // 
            this.上移ToolStripMenuItem.Name = "上移ToolStripMenuItem";
            this.上移ToolStripMenuItem.Size = new System.Drawing.Size(153, 22);
            this.上移ToolStripMenuItem.Text = "上移";
            this.上移ToolStripMenuItem.Click += new System.EventHandler(this.上移ToolStripMenuItem_Click);
            // 
            // 下移ToolStripMenuItem
            // 
            this.下移ToolStripMenuItem.Name = "下移ToolStripMenuItem";
            this.下移ToolStripMenuItem.Size = new System.Drawing.Size(153, 22);
            this.下移ToolStripMenuItem.Text = "下移";
            this.下移ToolStripMenuItem.Click += new System.EventHandler(this.下移ToolStripMenuItem_Click);
            // 
            // bindingSource1
            // 
            this.bindingSource1.DataSource = typeof(System.Collections.Generic.List<LTZN.CalFramework.CalTag>);
            // 
            // flowLayoutPanel1
            // 
            this.flowLayoutPanel1.Controls.Add(this.ckbViewType);
            this.flowLayoutPanel1.Controls.Add(this.label1);
            this.flowLayoutPanel1.Controls.Add(this.cmbFilterTyp);
            this.flowLayoutPanel1.Controls.Add(this.txtFilter);
            this.flowLayoutPanel1.Controls.Add(this.buttonCal);
            this.flowLayoutPanel1.Controls.Add(this.buttonOk);
            this.flowLayoutPanel1.Controls.Add(this.buttonCancel);
            this.flowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.flowLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.flowLayoutPanel1.Name = "flowLayoutPanel1";
            this.flowLayoutPanel1.Size = new System.Drawing.Size(886, 36);
            this.flowLayoutPanel1.TabIndex = 13;
            // 
            // ckbViewType
            // 
            this.ckbViewType.AutoSize = true;
            this.ckbViewType.Location = new System.Drawing.Point(8, 8);
            this.ckbViewType.Margin = new System.Windows.Forms.Padding(8);
            this.ckbViewType.Name = "ckbViewType";
            this.ckbViewType.Size = new System.Drawing.Size(96, 16);
            this.ckbViewType.TabIndex = 15;
            this.ckbViewType.Text = "显示所有标签";
            this.ckbViewType.UseVisualStyleBackColor = true;
            this.ckbViewType.CheckedChanged += new System.EventHandler(this.ckbViewType_CheckedChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(130, 8);
            this.label1.Margin = new System.Windows.Forms.Padding(18, 8, 8, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(29, 12);
            this.label1.TabIndex = 14;
            this.label1.Text = "筛选";
            // 
            // cmbFilterTyp
            // 
            this.cmbFilterTyp.FormattingEnabled = true;
            this.cmbFilterTyp.Items.AddRange(new object[] {
            "标签全名",
            "标签名",
            "公式"});
            this.cmbFilterTyp.Location = new System.Drawing.Point(170, 6);
            this.cmbFilterTyp.Margin = new System.Windows.Forms.Padding(3, 6, 3, 3);
            this.cmbFilterTyp.Name = "cmbFilterTyp";
            this.cmbFilterTyp.Size = new System.Drawing.Size(74, 20);
            this.cmbFilterTyp.TabIndex = 16;
            this.cmbFilterTyp.Text = "标签全名";
            // 
            // txtFilter
            // 
            this.txtFilter.Location = new System.Drawing.Point(252, 5);
            this.txtFilter.Margin = new System.Windows.Forms.Padding(5);
            this.txtFilter.Name = "txtFilter";
            this.txtFilter.Size = new System.Drawing.Size(209, 21);
            this.txtFilter.TabIndex = 13;
            this.txtFilter.TextChanged += new System.EventHandler(this.txtFilter_TextChanged);
            // 
            // buttonOk
            // 
            this.buttonOk.Location = new System.Drawing.Point(667, 3);
            this.buttonOk.Name = "buttonOk";
            this.buttonOk.Size = new System.Drawing.Size(75, 23);
            this.buttonOk.TabIndex = 17;
            this.buttonOk.Text = "保存";
            this.buttonOk.UseVisualStyleBackColor = true;
            this.buttonOk.Click += new System.EventHandler(this.buttonOk_Click);
            // 
            // buttonCal
            // 
            this.buttonCal.Location = new System.Drawing.Point(586, 3);
            this.buttonCal.Margin = new System.Windows.Forms.Padding(120, 3, 3, 3);
            this.buttonCal.Name = "buttonCal";
            this.buttonCal.Size = new System.Drawing.Size(75, 23);
            this.buttonCal.TabIndex = 18;
            this.buttonCal.Text = "计算";
            this.buttonCal.UseVisualStyleBackColor = true;
            this.buttonCal.Click += new System.EventHandler(this.计算ToolStripMenuItem_Click);
            // 
            // buttonCancel
            // 
            this.buttonCancel.Location = new System.Drawing.Point(748, 3);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 19;
            this.buttonCancel.Text = "取消";
            this.buttonCancel.UseVisualStyleBackColor = true;
            this.buttonCancel.Click += new System.EventHandler(this.buttonCancel_Click);
            // 
            // CalModelForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1094, 445);
            this.Controls.Add(this.splitContainer1);
            this.Name = "CalModelForm";
            this.Text = "计算模型";
            this.Load += new System.EventHandler(this.CalModelForm_Load);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.ResumeLayout(false);
            this.contextMenuStrip1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.c1FlexGrid1)).EndInit();
            this.contextMenuStrip2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            this.flowLayoutPanel1.ResumeLayout(false);
            this.flowLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.TreeView treeView1;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem btnAddModel;
        private System.Windows.Forms.ToolStripMenuItem btnAddOrg;
        private System.Windows.Forms.ToolStripMenuItem btnAddTag;
        private System.Windows.Forms.ToolStripMenuItem btnSave;
        private System.Windows.Forms.ToolStripMenuItem btnDeleteModel;
        private System.Windows.Forms.ToolStripMenuItem btnDeleteOrg;
        private System.Windows.Forms.BindingSource bindingSource1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip2;
        private System.Windows.Forms.ToolStripMenuItem btnAddTag2;
        private System.Windows.Forms.ToolStripMenuItem btnDeleteTag;
        private System.Windows.Forms.ToolStripMenuItem btnCopy;
        private System.Windows.Forms.ToolStripMenuItem btnPaste;
        private System.Windows.Forms.ToolStripMenuItem btnSumForma;
        private System.Windows.Forms.ToolStripMenuItem btnFenTanForma;
        private System.Windows.Forms.ToolStripMenuItem btnCalc;
        private System.Windows.Forms.FlowLayoutPanel flowLayoutPanel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtFilter;
        private C1.Win.C1FlexGrid.C1FlexGrid c1FlexGrid1;
        private System.Windows.Forms.CheckBox ckbViewType;
        private System.Windows.Forms.ToolStripMenuItem 设定累计公式ToolStripMenuItem;
        private System.Windows.Forms.ComboBox cmbFilterTyp;
        private System.Windows.Forms.ToolStripMenuItem 上移ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 下移ToolStripMenuItem;
        private System.Windows.Forms.Button buttonCal;
        private System.Windows.Forms.Button buttonOk;
        private System.Windows.Forms.Button buttonCancel;

    }
}