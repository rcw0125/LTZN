namespace LTZN.全厂日消耗
{
    partial class 全厂日消耗窗体
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(全厂日消耗窗体));
            this.panel1 = new System.Windows.Forms.Panel();
            this.flex原料 = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.bindingSource2 = new System.Windows.Forms.BindingSource(this.components);
            this.全厂日消耗数据集1 = new LTZN.全厂日消耗.全厂日消耗数据集();
            this.c1TextBox2 = new C1.Win.C1Input.C1TextBox();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.c1TextBox1 = new C1.Win.C1Input.C1TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.c1NumericEdit7 = new C1.Win.C1Input.C1NumericEdit();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.c1NumericEdit6 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit5 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit4 = new C1.Win.C1Input.C1NumericEdit();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.c1NumericEdit3 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit2 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit1 = new C1.Win.C1Input.C1NumericEdit();
            this.button1 = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.全厂日消耗TableAdapter = new LTZN.全厂日消耗.全厂日消耗数据集TableAdapters.全厂日消耗TableAdapter();
            this.日原料成分TableAdapter = new LTZN.全厂日消耗.全厂日消耗数据集TableAdapters.日原料成分TableAdapter();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.flex原料)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.全厂日消耗数据集1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1TextBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1TextBox1)).BeginInit();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit7)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit6)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit4)).BeginInit();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.flex原料);
            this.panel1.Controls.Add(this.c1TextBox2);
            this.panel1.Controls.Add(this.c1TextBox1);
            this.panel1.Controls.Add(this.groupBox2);
            this.panel1.Controls.Add(this.label10);
            this.panel1.Controls.Add(this.label9);
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.dateTimePicker1);
            this.panel1.Location = new System.Drawing.Point(6, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(729, 504);
            this.panel1.TabIndex = 0;
            // 
            // flex原料
            // 
            this.flex原料.AutoGenerateColumns = false;
            this.flex原料.ColumnInfo = resources.GetString("flex原料.ColumnInfo");
            this.flex原料.DataSource = this.bindingSource2;
            this.flex原料.Font = new System.Drawing.Font("宋体", 12F);
            this.flex原料.KeyActionEnter = C1.Win.C1FlexGrid.KeyActionEnum.MoveAcross;
            this.flex原料.Location = new System.Drawing.Point(226, 80);
            this.flex原料.Name = "flex原料";
            this.flex原料.Rows.Count = 1;
            this.flex原料.Rows.DefaultSize = 23;
            this.flex原料.SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.Row;
            this.flex原料.Size = new System.Drawing.Size(477, 321);
            this.flex原料.StyleInfo = resources.GetString("flex原料.StyleInfo");
            this.flex原料.TabIndex = 6;
            this.flex原料.Tag = "";
            // 
            // bindingSource2
            // 
            this.bindingSource2.DataMember = "日原料成分";
            this.bindingSource2.DataSource = this.全厂日消耗数据集1;
            // 
            // 全厂日消耗数据集1
            // 
            this.全厂日消耗数据集1.DataSetName = "全厂日消耗数据集";
            this.全厂日消耗数据集1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // c1TextBox2
            // 
            this.c1TextBox2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.bindingSource1, "P10备注2", true));
            this.c1TextBox2.EmptyAsNull = true;
            this.c1TextBox2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.c1TextBox2.Location = new System.Drawing.Point(96, 454);
            this.c1TextBox2.Name = "c1TextBox2";
            this.c1TextBox2.Size = new System.Drawing.Size(607, 26);
            this.c1TextBox2.TabIndex = 5;
            this.c1TextBox2.Tag = null;
            this.c1TextBox2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1TextBox1_KeyDown);
            // 
            // bindingSource1
            // 
            this.bindingSource1.DataMember = "全厂日消耗";
            this.bindingSource1.DataSource = this.全厂日消耗数据集1;
            // 
            // c1TextBox1
            // 
            this.c1TextBox1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.bindingSource1, "P09备注1", true));
            this.c1TextBox1.EmptyAsNull = true;
            this.c1TextBox1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.c1TextBox1.Location = new System.Drawing.Point(96, 422);
            this.c1TextBox1.Name = "c1TextBox1";
            this.c1TextBox1.Size = new System.Drawing.Size(607, 26);
            this.c1TextBox1.TabIndex = 4;
            this.c1TextBox1.Tag = null;
            this.c1TextBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1TextBox1_KeyDown);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.c1NumericEdit7);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.c1NumericEdit6);
            this.groupBox2.Controls.Add(this.c1NumericEdit5);
            this.groupBox2.Controls.Add(this.c1NumericEdit4);
            this.groupBox2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox2.Location = new System.Drawing.Point(20, 241);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(200, 160);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            // 
            // c1NumericEdit7
            // 
            this.c1NumericEdit7.Culture = 127;
            this.c1NumericEdit7.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.bindingSource1, "P08烧结二号称", true));
            this.c1NumericEdit7.DataType = typeof(double);
            this.c1NumericEdit7.EmptyAsNull = true;
            this.c1NumericEdit7.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.c1NumericEdit7.Location = new System.Drawing.Point(90, 128);
            this.c1NumericEdit7.Name = "c1NumericEdit7";
            this.c1NumericEdit7.Size = new System.Drawing.Size(94, 22);
            this.c1NumericEdit7.TabIndex = 4;
            this.c1NumericEdit7.Tag = null;
            this.c1NumericEdit7.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit7.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1NumericEdit27_KeyDown);
            // 
            // label8
            // 
            this.label8.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label8.Location = new System.Drawing.Point(6, 128);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(72, 24);
            this.label8.TabIndex = 44;
            this.label8.Text = "烧结2#称";
            // 
            // label7
            // 
            this.label7.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label7.Location = new System.Drawing.Point(6, 89);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(64, 24);
            this.label7.TabIndex = 43;
            this.label7.Text = "反矿总量";
            // 
            // label6
            // 
            this.label6.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label6.Location = new System.Drawing.Point(6, 56);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(64, 24);
            this.label6.TabIndex = 42;
            this.label6.Text = "3#皮带数";
            // 
            // label5
            // 
            this.label5.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.Location = new System.Drawing.Point(6, 22);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(64, 24);
            this.label5.TabIndex = 41;
            this.label5.Text = "2#皮带数";
            // 
            // c1NumericEdit6
            // 
            this.c1NumericEdit6.Culture = 127;
            this.c1NumericEdit6.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.bindingSource1, "P07总返矿", true));
            this.c1NumericEdit6.DataType = typeof(double);
            this.c1NumericEdit6.EmptyAsNull = true;
            this.c1NumericEdit6.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.c1NumericEdit6.Location = new System.Drawing.Point(90, 91);
            this.c1NumericEdit6.Name = "c1NumericEdit6";
            this.c1NumericEdit6.Size = new System.Drawing.Size(94, 22);
            this.c1NumericEdit6.TabIndex = 3;
            this.c1NumericEdit6.Tag = null;
            this.c1NumericEdit6.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit6.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1NumericEdit27_KeyDown);
            // 
            // c1NumericEdit5
            // 
            this.c1NumericEdit5.Culture = 127;
            this.c1NumericEdit5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.bindingSource1, "P06三号皮带", true));
            this.c1NumericEdit5.DataType = typeof(double);
            this.c1NumericEdit5.EmptyAsNull = true;
            this.c1NumericEdit5.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.c1NumericEdit5.Location = new System.Drawing.Point(90, 56);
            this.c1NumericEdit5.Name = "c1NumericEdit5";
            this.c1NumericEdit5.Size = new System.Drawing.Size(94, 22);
            this.c1NumericEdit5.TabIndex = 2;
            this.c1NumericEdit5.Tag = null;
            this.c1NumericEdit5.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit5.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1NumericEdit27_KeyDown);
            // 
            // c1NumericEdit4
            // 
            this.c1NumericEdit4.Culture = 127;
            this.c1NumericEdit4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.bindingSource1, "P05二号皮带", true));
            this.c1NumericEdit4.DataType = typeof(double);
            this.c1NumericEdit4.EmptyAsNull = true;
            this.c1NumericEdit4.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.c1NumericEdit4.Location = new System.Drawing.Point(90, 18);
            this.c1NumericEdit4.Name = "c1NumericEdit4";
            this.c1NumericEdit4.Size = new System.Drawing.Size(94, 22);
            this.c1NumericEdit4.TabIndex = 1;
            this.c1NumericEdit4.Tag = null;
            this.c1NumericEdit4.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit4.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1NumericEdit27_KeyDown);
            // 
            // label10
            // 
            this.label10.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label10.Location = new System.Drawing.Point(26, 456);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(64, 24);
            this.label10.TabIndex = 26;
            this.label10.Text = "备注2";
            // 
            // label9
            // 
            this.label9.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label9.Location = new System.Drawing.Point(26, 424);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(64, 24);
            this.label9.TabIndex = 25;
            this.label9.Text = "备注1";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.c1NumericEdit3);
            this.groupBox1.Controls.Add(this.c1NumericEdit2);
            this.groupBox1.Controls.Add(this.c1NumericEdit1);
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox1.Location = new System.Drawing.Point(20, 71);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(200, 168);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "全天消耗输入";
            // 
            // c1NumericEdit3
            // 
            this.c1NumericEdit3.Culture = 127;
            this.c1NumericEdit3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.bindingSource1, "P04焦粉水分", true));
            this.c1NumericEdit3.DataType = typeof(double);
            this.c1NumericEdit3.EmptyAsNull = true;
            this.c1NumericEdit3.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.c1NumericEdit3.Location = new System.Drawing.Point(117, 91);
            this.c1NumericEdit3.Name = "c1NumericEdit3";
            this.c1NumericEdit3.Size = new System.Drawing.Size(67, 22);
            this.c1NumericEdit3.TabIndex = 3;
            this.c1NumericEdit3.Tag = null;
            this.c1NumericEdit3.Visible = false;
            this.c1NumericEdit3.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit3.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1NumericEdit27_KeyDown);
            // 
            // c1NumericEdit2
            // 
            this.c1NumericEdit2.Culture = 127;
            this.c1NumericEdit2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.bindingSource1, "P03落地焦水分", true));
            this.c1NumericEdit2.DataType = typeof(double);
            this.c1NumericEdit2.EmptyAsNull = true;
            this.c1NumericEdit2.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.c1NumericEdit2.Location = new System.Drawing.Point(117, 58);
            this.c1NumericEdit2.Name = "c1NumericEdit2";
            this.c1NumericEdit2.Size = new System.Drawing.Size(67, 22);
            this.c1NumericEdit2.TabIndex = 2;
            this.c1NumericEdit2.Tag = null;
            this.c1NumericEdit2.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1NumericEdit27_KeyDown);
            // 
            // c1NumericEdit1
            // 
            this.c1NumericEdit1.Culture = 127;
            this.c1NumericEdit1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.bindingSource1, "P02自产焦水分", true));
            this.c1NumericEdit1.DataType = typeof(double);
            this.c1NumericEdit1.EmptyAsNull = true;
            this.c1NumericEdit1.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.c1NumericEdit1.Location = new System.Drawing.Point(117, 25);
            this.c1NumericEdit1.Name = "c1NumericEdit1";
            this.c1NumericEdit1.Size = new System.Drawing.Size(67, 22);
            this.c1NumericEdit1.TabIndex = 1;
            this.c1NumericEdit1.Tag = null;
            this.c1NumericEdit1.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1NumericEdit27_KeyDown);
            // 
            // button1
            // 
            this.button1.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
            this.button1.Location = new System.Drawing.Point(32, 128);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(128, 32);
            this.button1.TabIndex = 6;
            this.button1.Text = "提取焦炭数据";
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label4
            // 
            this.label4.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(16, 96);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(96, 24);
            this.label4.TabIndex = 2;
            this.label4.Text = "焦粉  水分";
            this.label4.Visible = false;
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(15, 64);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(96, 24);
            this.label3.TabIndex = 1;
            this.label3.Text = "落地焦水分";
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(16, 32);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(96, 24);
            this.label2.TabIndex = 0;
            this.label2.Text = "自产焦水分";
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 21.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.ForeColor = System.Drawing.Color.Lime;
            this.label1.Location = new System.Drawing.Point(275, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(232, 40);
            this.label1.TabIndex = 23;
            this.label1.Text = "全厂日销耗录入";
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dateTimePicker1.Location = new System.Drawing.Point(20, 36);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(139, 26);
            this.dateTimePicker1.TabIndex = 1;
            this.dateTimePicker1.ValueChanged += new System.EventHandler(this.dateTimePicker1_ValueChanged);
            // 
            // 全厂日消耗TableAdapter
            // 
            this.全厂日消耗TableAdapter.ClearBeforeFill = true;
            // 
            // 日原料成分TableAdapter
            // 
            this.日原料成分TableAdapter.ClearBeforeFill = true;
            // 
            // 全厂日消耗窗体
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSeaGreen;
            this.ClientSize = new System.Drawing.Size(768, 535);
            this.Controls.Add(this.panel1);
            this.Name = "全厂日消耗窗体";
            this.Text = "全厂日消耗窗体";
            this.Load += new System.EventHandler(this.全厂日消耗窗体_Load);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.全厂日消耗窗体_FormClosed);
            this.Leave += new System.EventHandler(this.全厂日消耗窗体_Leave);
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.flex原料)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.全厂日消耗数据集1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1TextBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1TextBox1)).EndInit();
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit7)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit6)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit4)).EndInit();
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private C1.Win.C1FlexGrid.C1FlexGrid flex原料;
        private C1.Win.C1Input.C1TextBox c1TextBox2;
        private C1.Win.C1Input.C1TextBox c1TextBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit6;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit5;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit4;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.GroupBox groupBox1;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit3;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit2;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private 全厂日消耗数据集 全厂日消耗数据集1;
        private System.Windows.Forms.BindingSource bindingSource1;
        private System.Windows.Forms.BindingSource bindingSource2;
        private LTZN.全厂日消耗.全厂日消耗数据集TableAdapters.全厂日消耗TableAdapter 全厂日消耗TableAdapter;
        private LTZN.全厂日消耗.全厂日消耗数据集TableAdapters.日原料成分TableAdapter 日原料成分TableAdapter;
    }
}