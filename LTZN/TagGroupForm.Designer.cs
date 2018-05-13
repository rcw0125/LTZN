namespace LTZN
{
    partial class TagGroupForm
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
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.lstTagGroups = new System.Windows.Forms.ListBox();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.添加标签组ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.删除标签组ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label1 = new System.Windows.Forms.Label();
            this.lstTags = new System.Windows.Forms.ListBox();
            this.contextMenuStrip2 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.添加标签ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.删除标签ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.label2 = new System.Windows.Forms.Label();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.contextMenuStrip1.SuspendLayout();
            this.contextMenuStrip2.SuspendLayout();
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
            this.splitContainer1.Panel1.Controls.Add(this.lstTagGroups);
            this.splitContainer1.Panel1.Controls.Add(this.label1);
            this.splitContainer1.Panel1.Padding = new System.Windows.Forms.Padding(3);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.lstTags);
            this.splitContainer1.Panel2.Controls.Add(this.label2);
            this.splitContainer1.Panel2.Padding = new System.Windows.Forms.Padding(5);
            this.splitContainer1.Size = new System.Drawing.Size(444, 395);
            this.splitContainer1.SplitterDistance = 148;
            this.splitContainer1.TabIndex = 0;
            // 
            // lstTagGroups
            // 
            this.lstTagGroups.ContextMenuStrip = this.contextMenuStrip1;
            this.lstTagGroups.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lstTagGroups.FormattingEnabled = true;
            this.lstTagGroups.ItemHeight = 12;
            this.lstTagGroups.Location = new System.Drawing.Point(3, 25);
            this.lstTagGroups.Name = "lstTagGroups";
            this.lstTagGroups.Size = new System.Drawing.Size(142, 364);
            this.lstTagGroups.TabIndex = 2;
            this.lstTagGroups.SelectedIndexChanged += new System.EventHandler(this.lstTagGroups_SelectedIndexChanged);
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.添加标签组ToolStripMenuItem,
            this.删除标签组ToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(131, 48);
            // 
            // 添加标签组ToolStripMenuItem
            // 
            this.添加标签组ToolStripMenuItem.Name = "添加标签组ToolStripMenuItem";
            this.添加标签组ToolStripMenuItem.Size = new System.Drawing.Size(130, 22);
            this.添加标签组ToolStripMenuItem.Text = "添加标签组";
            this.添加标签组ToolStripMenuItem.Click += new System.EventHandler(this.添加标签组ToolStripMenuItem_Click);
            // 
            // 删除标签组ToolStripMenuItem
            // 
            this.删除标签组ToolStripMenuItem.Name = "删除标签组ToolStripMenuItem";
            this.删除标签组ToolStripMenuItem.Size = new System.Drawing.Size(130, 22);
            this.删除标签组ToolStripMenuItem.Text = "删除标签组";
            this.删除标签组ToolStripMenuItem.Click += new System.EventHandler(this.删除标签组ToolStripMenuItem_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Dock = System.Windows.Forms.DockStyle.Top;
            this.label1.Location = new System.Drawing.Point(3, 3);
            this.label1.Name = "label1";
            this.label1.Padding = new System.Windows.Forms.Padding(5);
            this.label1.Size = new System.Drawing.Size(51, 22);
            this.label1.TabIndex = 1;
            this.label1.Text = "标签组";
            // 
            // lstTags
            // 
            this.lstTags.ContextMenuStrip = this.contextMenuStrip2;
            this.lstTags.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lstTags.FormattingEnabled = true;
            this.lstTags.ItemHeight = 12;
            this.lstTags.Location = new System.Drawing.Point(5, 27);
            this.lstTags.Name = "lstTags";
            this.lstTags.Size = new System.Drawing.Size(282, 352);
            this.lstTags.TabIndex = 3;
            // 
            // contextMenuStrip2
            // 
            this.contextMenuStrip2.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.添加标签ToolStripMenuItem,
            this.删除标签ToolStripMenuItem});
            this.contextMenuStrip2.Name = "contextMenuStrip2";
            this.contextMenuStrip2.Size = new System.Drawing.Size(119, 48);
            // 
            // 添加标签ToolStripMenuItem
            // 
            this.添加标签ToolStripMenuItem.Name = "添加标签ToolStripMenuItem";
            this.添加标签ToolStripMenuItem.Size = new System.Drawing.Size(118, 22);
            this.添加标签ToolStripMenuItem.Text = "添加标签";
            this.添加标签ToolStripMenuItem.Click += new System.EventHandler(this.添加标签ToolStripMenuItem_Click);
            // 
            // 删除标签ToolStripMenuItem
            // 
            this.删除标签ToolStripMenuItem.Name = "删除标签ToolStripMenuItem";
            this.删除标签ToolStripMenuItem.Size = new System.Drawing.Size(118, 22);
            this.删除标签ToolStripMenuItem.Text = "删除标签";
            this.删除标签ToolStripMenuItem.Click += new System.EventHandler(this.删除标签ToolStripMenuItem_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Dock = System.Windows.Forms.DockStyle.Top;
            this.label2.Location = new System.Drawing.Point(5, 5);
            this.label2.Name = "label2";
            this.label2.Padding = new System.Windows.Forms.Padding(5);
            this.label2.Size = new System.Drawing.Size(39, 22);
            this.label2.TabIndex = 2;
            this.label2.Text = "标签";
            // 
            // TagGroupForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(444, 395);
            this.Controls.Add(this.splitContainer1);
            this.Name = "TagGroupForm";
            this.Text = "标签组";
            this.Load += new System.EventHandler(this.TagGroupForm_Load);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.TagGroupForm_FormClosed);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel1.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.Panel2.PerformLayout();
            this.splitContainer1.ResumeLayout(false);
            this.contextMenuStrip1.ResumeLayout(false);
            this.contextMenuStrip2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox lstTagGroups;
        private System.Windows.Forms.ListBox lstTags;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip2;
        private System.Windows.Forms.ToolStripMenuItem 添加标签组ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 删除标签组ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 添加标签ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 删除标签ToolStripMenuItem;
    }
}