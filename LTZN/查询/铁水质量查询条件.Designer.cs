namespace LTZN.查询
{
    partial class 铁水质量查询条件
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
            this.button1 = new System.Windows.Forms.Button();
            this.TICheck = new System.Windows.Forms.CheckBox();
            this.PCheck = new System.Windows.Forms.CheckBox();
            this.SCheck = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button1.Location = new System.Drawing.Point(165, 179);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(76, 24);
            this.button1.TabIndex = 3;
            this.button1.Text = "完成";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // TICheck
            // 
            this.TICheck.AutoSize = true;
            this.TICheck.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.TICheck.Location = new System.Drawing.Point(63, 105);
            this.TICheck.Name = "TICheck";
            this.TICheck.Size = new System.Drawing.Size(142, 20);
            this.TICheck.TabIndex = 8;
            this.TICheck.Text = "铁水TI<=0.055";
            this.TICheck.UseVisualStyleBackColor = true;
            // 
            // PCheck
            // 
            this.PCheck.AutoSize = true;
            this.PCheck.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.PCheck.Location = new System.Drawing.Point(63, 74);
            this.PCheck.Name = "PCheck";
            this.PCheck.Size = new System.Drawing.Size(133, 20);
            this.PCheck.TabIndex = 7;
            this.PCheck.Text = "铁水P<=0.009";
            this.PCheck.UseVisualStyleBackColor = true;
            // 
            // SCheck
            // 
            this.SCheck.AutoSize = true;
            this.SCheck.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.SCheck.Location = new System.Drawing.Point(63, 43);
            this.SCheck.Name = "SCheck";
            this.SCheck.Size = new System.Drawing.Size(124, 20);
            this.SCheck.TabIndex = 6;
            this.SCheck.Text = "铁水S<=0.02";
            this.SCheck.UseVisualStyleBackColor = true;
            // 
            // 铁水质量查询条件
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(288, 216);
            this.Controls.Add(this.TICheck);
            this.Controls.Add(this.PCheck);
            this.Controls.Add(this.SCheck);
            this.Controls.Add(this.button1);
            this.Name = "铁水质量查询条件";
            this.Text = "请选择查询条件";
            this.Load += new System.EventHandler(this.铁水质量查询条件_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.CheckBox TICheck;
        private System.Windows.Forms.CheckBox PCheck;
        private System.Windows.Forms.CheckBox SCheck;
    }
}