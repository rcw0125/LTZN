using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using C1.Win.C1FlexGrid;
using System.Data.OracleClient;
using C1.Win.C1Input;

namespace LTZN.XiaohaoGaolu
{
	/// <summary>
	/// Summary description for Form2.
	/// </summary>
	public class XiaohaoGaolu : System.Windows.Forms.Form
    {
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.ComboBox comboBox1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.GroupBox groupBox6;
        private System.Windows.Forms.GroupBox groupBox7;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit0;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit11;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit10;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit9;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit8;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit6;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit5;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit4;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit3;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit2;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit1;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit12;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit25;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit24;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit23;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit30;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit29;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit28;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit27;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit26;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit22;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit21;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit20;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit19;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit18;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit17;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit16;
        private GroupBox groupBox2;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit15;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit14;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit13;
        private BindingSource xiaohaoManagerBindingSource;
        private BindingSource bindingSource1;
        private DataSetXh1 dataSetXh1;
        private C1.Win.C1Command.C1DockingTab c1DockingTab1;
        private C1.Win.C1Command.C1DockingTabPage 休风;
        private C1FlexGrid c1FlexGrid1;
        private C1.Win.C1Command.C1DockingTabPage 慢风;
        private BindingSource bindingSource2;
        private LTZN.XiaohaoGaolu.DataSetXh1TableAdapters.慢风TableAdapter 慢风TableAdapter;
        private LTZN.XiaohaoGaolu.DataSetXh1TableAdapters.休风TableAdapter 休风TableAdapter;
        private Label label1;
        private Label label2;
        private Label label5;
        private Label label4;
        private Label label3;
        private GroupBox groupBox5;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit31;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit32;
        private C1.Win.C1Input.C1NumericEdit c1NumericEdit33;
        private Button button1;
        private Label label9;
        private Label label8;
        private Label label7;
        private Label label6;
        private Label label15;
        private Label label14;
        private Label label13;
        private Label label12;
        private Label label11;
        private Label label10;
        private Label label29;
        private Label label28;
        private Label label27;
        private Label label18;
        private Label label19;
        private Label label30;
        private Label label31;
        private Label label32;
        private Label label26;
        private Label label25;
        private Label label24;
        private Label label23;
        private Label label22;
        private Label label20;
        private Label label21;
        private Label label17;
        private Label label16;
        private Label label33;
        private Label label34;
        private Label label35;
        private C1FlexGrid c1FlexGrid2;
        private IContainer components;
        private C1NumericEdit c1NumericEdit34;
        private Label label36;
        private C1NumericEdit c1NumericEdit7;

        private bool canUpdate = false;
        public XiaohaoGaolu()
        {
            //
            // Required for Windows Form Designer support
            //
            InitializeComponent();

            c1FlexGrid1.AllowAddNew = false;
            c1FlexGrid1.AllowDelete = false;
            c1FlexGrid1.AllowEditing = false;

            c1FlexGrid2.AllowAddNew = false;
            c1FlexGrid2.AllowDelete = false;
            c1FlexGrid2.AllowEditing = false;

            LtznUserManager.instance.RegisterHandler(this, XiaohaoGaolu_UserChange);
        }

        void XiaohaoGaolu_UserChange(LtznUser ltznUser)
        {
            if (ltznUser!=null && ltznUser.IsInRole("统计"))
            {
                canUpdate = true;
                c1FlexGrid1.AllowAddNew = true;
                c1FlexGrid1.AllowDelete = true;
                c1FlexGrid1.AllowEditing = true;

                c1FlexGrid2.AllowAddNew = true;
                c1FlexGrid2.AllowDelete = true;
                c1FlexGrid2.AllowEditing = true;

                button1.Enabled = true;
            }
            else
            {
                canUpdate = false;
                c1FlexGrid1.AllowAddNew = false;
                c1FlexGrid1.AllowDelete = false;
                c1FlexGrid1.AllowEditing = false;

                c1FlexGrid2.AllowAddNew = false;
                c1FlexGrid2.AllowDelete = false;
                c1FlexGrid2.AllowEditing = false;

                button1.Enabled = false;
            }
        }

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(XiaohaoGaolu));
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.c1NumericEdit34 = new C1.Win.C1Input.C1NumericEdit();
            this.label15 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.c1NumericEdit12 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit11 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit10 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit9 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit8 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit6 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit5 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit4 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit3 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit2 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit1 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit0 = new C1.Win.C1Input.C1NumericEdit();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.label29 = new System.Windows.Forms.Label();
            this.label28 = new System.Windows.Forms.Label();
            this.label27 = new System.Windows.Forms.Label();
            this.c1NumericEdit25 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit24 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit23 = new C1.Win.C1Input.C1NumericEdit();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.label18 = new System.Windows.Forms.Label();
            this.label19 = new System.Windows.Forms.Label();
            this.c1NumericEdit30 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit29 = new C1.Win.C1Input.C1NumericEdit();
            this.groupBox6 = new System.Windows.Forms.GroupBox();
            this.label30 = new System.Windows.Forms.Label();
            this.label31 = new System.Windows.Forms.Label();
            this.label32 = new System.Windows.Forms.Label();
            this.c1NumericEdit28 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit27 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit26 = new C1.Win.C1Input.C1NumericEdit();
            this.groupBox7 = new System.Windows.Forms.GroupBox();
            this.label26 = new System.Windows.Forms.Label();
            this.label25 = new System.Windows.Forms.Label();
            this.label24 = new System.Windows.Forms.Label();
            this.label23 = new System.Windows.Forms.Label();
            this.label22 = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.label21 = new System.Windows.Forms.Label();
            this.c1NumericEdit22 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit21 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit20 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit19 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit18 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit17 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit16 = new C1.Win.C1Input.C1NumericEdit();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label17 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.c1NumericEdit15 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit14 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit13 = new C1.Win.C1Input.C1NumericEdit();
            this.c1DockingTab1 = new C1.Win.C1Command.C1DockingTab();
            this.休风 = new C1.Win.C1Command.C1DockingTabPage();
            this.c1FlexGrid1 = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.dataSetXh1 = new LTZN.XiaohaoGaolu.DataSetXh1();
            this.慢风 = new C1.Win.C1Command.C1DockingTabPage();
            this.c1FlexGrid2 = new C1.Win.C1FlexGrid.C1FlexGrid();
            this.bindingSource2 = new System.Windows.Forms.BindingSource(this.components);
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.label33 = new System.Windows.Forms.Label();
            this.label34 = new System.Windows.Forms.Label();
            this.label35 = new System.Windows.Forms.Label();
            this.c1NumericEdit31 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit32 = new C1.Win.C1Input.C1NumericEdit();
            this.c1NumericEdit33 = new C1.Win.C1Input.C1NumericEdit();
            this.button1 = new System.Windows.Forms.Button();
            this.慢风TableAdapter = new LTZN.XiaohaoGaolu.DataSetXh1TableAdapters.慢风TableAdapter();
            this.休风TableAdapter = new LTZN.XiaohaoGaolu.DataSetXh1TableAdapters.休风TableAdapter();
            this.label36 = new System.Windows.Forms.Label();
            this.c1NumericEdit7 = new C1.Win.C1Input.C1NumericEdit();
            this.xiaohaoManagerBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit34)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit12)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit11)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit10)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit9)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit8)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit6)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit5)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit4)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit3)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit0)).BeginInit();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit25)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit24)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit23)).BeginInit();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit30)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit29)).BeginInit();
            this.groupBox6.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit28)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit27)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit26)).BeginInit();
            this.groupBox7.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit22)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit21)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit20)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit19)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit18)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit17)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit16)).BeginInit();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit15)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit14)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit13)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1DockingTab1)).BeginInit();
            this.c1DockingTab1.SuspendLayout();
            this.休风.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1FlexGrid1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSetXh1)).BeginInit();
            this.慢风.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1FlexGrid2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource2)).BeginInit();
            this.groupBox5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit31)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit32)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit33)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit7)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.xiaohaoManagerBindingSource)).BeginInit();
            this.SuspendLayout();
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dateTimePicker1.Location = new System.Drawing.Point(35, 51);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(168, 29);
            this.dateTimePicker1.TabIndex = 1;
            this.dateTimePicker1.ValueChanged += new System.EventHandler(this.dateTimePicker1_ValueChanged);
            // 
            // comboBox1
            // 
            this.comboBox1.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.comboBox1.Items.AddRange(new object[] {
            "1",
            "2",
            "3",
            "4",
            "5",
            "6"});
            this.comboBox1.Location = new System.Drawing.Point(225, 53);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(48, 27);
            this.comboBox1.TabIndex = 2;
            this.comboBox1.Text = "1";
            this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label36);
            this.groupBox1.Controls.Add(this.c1NumericEdit7);
            this.groupBox1.Controls.Add(this.c1NumericEdit34);
            this.groupBox1.Controls.Add(this.label15);
            this.groupBox1.Controls.Add(this.label14);
            this.groupBox1.Controls.Add(this.label13);
            this.groupBox1.Controls.Add(this.label12);
            this.groupBox1.Controls.Add(this.label11);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.c1NumericEdit12);
            this.groupBox1.Controls.Add(this.c1NumericEdit11);
            this.groupBox1.Controls.Add(this.c1NumericEdit10);
            this.groupBox1.Controls.Add(this.c1NumericEdit9);
            this.groupBox1.Controls.Add(this.c1NumericEdit8);
            this.groupBox1.Controls.Add(this.c1NumericEdit6);
            this.groupBox1.Controls.Add(this.c1NumericEdit5);
            this.groupBox1.Controls.Add(this.c1NumericEdit4);
            this.groupBox1.Controls.Add(this.c1NumericEdit3);
            this.groupBox1.Controls.Add(this.c1NumericEdit2);
            this.groupBox1.Controls.Add(this.c1NumericEdit1);
            this.groupBox1.Controls.Add(this.c1NumericEdit0);
            this.groupBox1.Location = new System.Drawing.Point(35, 86);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(192, 425);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            // 
            // c1NumericEdit34
            // 
            this.c1NumericEdit34.Culture = 127;
            this.c1NumericEdit34.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "Shijiaofen", true));
            this.c1NumericEdit34.DataType = typeof(double);
            this.c1NumericEdit34.EmptyAsNull = true;
            this.c1NumericEdit34.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.c1NumericEdit34.ImeMode = System.Windows.Forms.ImeMode.Off;
            this.c1NumericEdit34.Location = new System.Drawing.Point(110, 248);
            this.c1NumericEdit34.Name = "c1NumericEdit34";
            this.c1NumericEdit34.Size = new System.Drawing.Size(67, 22);
            this.c1NumericEdit34.TabIndex = 50;
            this.c1NumericEdit34.Tag = null;
            this.c1NumericEdit34.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label15.ForeColor = System.Drawing.Color.Blue;
            this.label15.Location = new System.Drawing.Point(5, 395);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(104, 16);
            this.label15.TabIndex = 49;
            this.label15.Text = "高压操作时间";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label14.ForeColor = System.Drawing.Color.Blue;
            this.label14.Location = new System.Drawing.Point(5, 366);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(24, 16);
            this.label14.TabIndex = 48;
            this.label14.Text = "CO";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label13.ForeColor = System.Drawing.Color.Blue;
            this.label13.Location = new System.Drawing.Point(5, 337);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(32, 16);
            this.label13.TabIndex = 47;
            this.label13.Text = "CO2";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label12.ForeColor = System.Drawing.Color.Blue;
            this.label12.Location = new System.Drawing.Point(5, 308);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(40, 16);
            this.label12.TabIndex = 46;
            this.label12.Text = "炉数";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label11.ForeColor = System.Drawing.Color.Blue;
            this.label11.Location = new System.Drawing.Point(5, 78);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(72, 16);
            this.label11.TabIndex = 45;
            this.label11.Text = "焦丁水份";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label10.ForeColor = System.Drawing.Color.Blue;
            this.label10.Location = new System.Drawing.Point(5, 251);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(56, 16);
            this.label10.TabIndex = 44;
            this.label10.Text = "湿焦粉";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label9.ForeColor = System.Drawing.Color.Blue;
            this.label9.Location = new System.Drawing.Point(5, 222);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(72, 16);
            this.label9.TabIndex = 43;
            this.label9.Text = "炉顶压力";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label8.ForeColor = System.Drawing.Color.Blue;
            this.label8.Location = new System.Drawing.Point(5, 193);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(72, 16);
            this.label8.TabIndex = 42;
            this.label8.Text = "炉顶温度";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label7.ForeColor = System.Drawing.Color.Blue;
            this.label7.Location = new System.Drawing.Point(5, 164);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(72, 16);
            this.label7.TabIndex = 41;
            this.label7.Text = "热风压力";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label6.ForeColor = System.Drawing.Color.Blue;
            this.label6.Location = new System.Drawing.Point(5, 135);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(40, 16);
            this.label6.TabIndex = 40;
            this.label6.Text = "风温";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.ForeColor = System.Drawing.Color.Blue;
            this.label5.Location = new System.Drawing.Point(5, 106);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(72, 16);
            this.label5.TabIndex = 39;
            this.label5.Text = "冷风流量";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.ForeColor = System.Drawing.Color.Blue;
            this.label4.Location = new System.Drawing.Point(5, 48);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(40, 16);
            this.label4.TabIndex = 39;
            this.label4.Text = "焦丁";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.ForeColor = System.Drawing.Color.Blue;
            this.label3.Location = new System.Drawing.Point(5, 19);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(56, 16);
            this.label3.TabIndex = 38;
            this.label3.Text = "富氧量";
            // 
            // c1NumericEdit12
            // 
            this.c1NumericEdit12.Culture = 127;
            this.c1NumericEdit12.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "Gaoyashijian", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit12.EmptyAsNull = true;
            this.c1NumericEdit12.Location = new System.Drawing.Point(110, 391);
            this.c1NumericEdit12.Name = "c1NumericEdit12";
            this.c1NumericEdit12.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit12.TabIndex = 37;
            this.c1NumericEdit12.Tag = null;
            this.c1NumericEdit12.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit12.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit11
            // 
            this.c1NumericEdit11.Culture = 127;
            this.c1NumericEdit11.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "CO", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit11.EmptyAsNull = true;
            this.c1NumericEdit11.Location = new System.Drawing.Point(110, 362);
            this.c1NumericEdit11.Name = "c1NumericEdit11";
            this.c1NumericEdit11.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit11.TabIndex = 36;
            this.c1NumericEdit11.Tag = null;
            this.c1NumericEdit11.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit11.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit10
            // 
            this.c1NumericEdit10.Culture = 127;
            this.c1NumericEdit10.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "CO2", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit10.EmptyAsNull = true;
            this.c1NumericEdit10.Location = new System.Drawing.Point(110, 333);
            this.c1NumericEdit10.Name = "c1NumericEdit10";
            this.c1NumericEdit10.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit10.TabIndex = 35;
            this.c1NumericEdit10.Tag = null;
            this.c1NumericEdit10.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit10.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit9
            // 
            this.c1NumericEdit9.Culture = 127;
            this.c1NumericEdit9.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "Lushu", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit9.EmptyAsNull = true;
            this.c1NumericEdit9.Location = new System.Drawing.Point(110, 304);
            this.c1NumericEdit9.Name = "c1NumericEdit9";
            this.c1NumericEdit9.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit9.TabIndex = 10;
            this.c1NumericEdit9.Tag = null;
            this.c1NumericEdit9.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit9.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit8
            // 
            this.c1NumericEdit8.Culture = 127;
            this.c1NumericEdit8.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "JiaodingShuiFen", true));
            this.c1NumericEdit8.EmptyAsNull = true;
            this.c1NumericEdit8.Location = new System.Drawing.Point(110, 74);
            this.c1NumericEdit8.Name = "c1NumericEdit8";
            this.c1NumericEdit8.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit8.TabIndex = 9;
            this.c1NumericEdit8.Tag = null;
            this.c1NumericEdit8.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit8.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit6
            // 
            this.c1NumericEdit6.Culture = 127;
            this.c1NumericEdit6.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "Ludingyali", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit6.EmptyAsNull = true;
            this.c1NumericEdit6.Location = new System.Drawing.Point(110, 218);
            this.c1NumericEdit6.Name = "c1NumericEdit6";
            this.c1NumericEdit6.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit6.TabIndex = 7;
            this.c1NumericEdit6.Tag = null;
            this.c1NumericEdit6.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit6.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit5
            // 
            this.c1NumericEdit5.Culture = 127;
            this.c1NumericEdit5.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "Ludingwendu", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit5.EmptyAsNull = true;
            this.c1NumericEdit5.Location = new System.Drawing.Point(110, 189);
            this.c1NumericEdit5.Name = "c1NumericEdit5";
            this.c1NumericEdit5.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit5.TabIndex = 6;
            this.c1NumericEdit5.Tag = null;
            this.c1NumericEdit5.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit5.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit4
            // 
            this.c1NumericEdit4.Culture = 127;
            this.c1NumericEdit4.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "Refengyali", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit4.EmptyAsNull = true;
            this.c1NumericEdit4.Location = new System.Drawing.Point(110, 161);
            this.c1NumericEdit4.Name = "c1NumericEdit4";
            this.c1NumericEdit4.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit4.TabIndex = 5;
            this.c1NumericEdit4.Tag = null;
            this.c1NumericEdit4.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit4.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit3
            // 
            this.c1NumericEdit3.Culture = 127;
            this.c1NumericEdit3.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "Fengwen", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit3.EmptyAsNull = true;
            this.c1NumericEdit3.Location = new System.Drawing.Point(110, 132);
            this.c1NumericEdit3.Name = "c1NumericEdit3";
            this.c1NumericEdit3.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit3.TabIndex = 4;
            this.c1NumericEdit3.Tag = null;
            this.c1NumericEdit3.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit3.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit2
            // 
            this.c1NumericEdit2.Culture = 127;
            this.c1NumericEdit2.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "Lengfengliuliang", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit2.EmptyAsNull = true;
            this.c1NumericEdit2.Location = new System.Drawing.Point(110, 103);
            this.c1NumericEdit2.Name = "c1NumericEdit2";
            this.c1NumericEdit2.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit2.TabIndex = 3;
            this.c1NumericEdit2.Tag = null;
            this.c1NumericEdit2.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit2.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit1
            // 
            this.c1NumericEdit1.Culture = 127;
            this.c1NumericEdit1.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "Jiaoding", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit1.EmptyAsNull = true;
            this.c1NumericEdit1.Location = new System.Drawing.Point(110, 45);
            this.c1NumericEdit1.Name = "c1NumericEdit1";
            this.c1NumericEdit1.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit1.TabIndex = 2;
            this.c1NumericEdit1.Tag = null;
            this.c1NumericEdit1.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit0
            // 
            this.c1NumericEdit0.Culture = 127;
            this.c1NumericEdit0.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "Fuyang", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit0.EmptyAsNull = true;
            this.c1NumericEdit0.Location = new System.Drawing.Point(110, 16);
            this.c1NumericEdit0.Name = "c1NumericEdit0";
            this.c1NumericEdit0.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit0.TabIndex = 1;
            this.c1NumericEdit0.Tag = null;
            this.c1NumericEdit0.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit0.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.label29);
            this.groupBox3.Controls.Add(this.label28);
            this.groupBox3.Controls.Add(this.label27);
            this.groupBox3.Controls.Add(this.c1NumericEdit25);
            this.groupBox3.Controls.Add(this.c1NumericEdit24);
            this.groupBox3.Controls.Add(this.c1NumericEdit23);
            this.groupBox3.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox3.ForeColor = System.Drawing.Color.Blue;
            this.groupBox3.Location = new System.Drawing.Point(248, 317);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(140, 98);
            this.groupBox3.TabIndex = 7;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "炉况";
            // 
            // label29
            // 
            this.label29.AutoSize = true;
            this.label29.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label29.ForeColor = System.Drawing.Color.Blue;
            this.label29.Location = new System.Drawing.Point(6, 74);
            this.label29.Name = "label29";
            this.label29.Size = new System.Drawing.Size(40, 16);
            this.label29.TabIndex = 53;
            this.label29.Text = "崩料";
            // 
            // label28
            // 
            this.label28.AutoSize = true;
            this.label28.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label28.ForeColor = System.Drawing.Color.Blue;
            this.label28.Location = new System.Drawing.Point(6, 48);
            this.label28.Name = "label28";
            this.label28.Size = new System.Drawing.Size(40, 16);
            this.label28.TabIndex = 52;
            this.label28.Text = "坐料";
            // 
            // label27
            // 
            this.label27.AutoSize = true;
            this.label27.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label27.ForeColor = System.Drawing.Color.Blue;
            this.label27.Location = new System.Drawing.Point(6, 22);
            this.label27.Name = "label27";
            this.label27.Size = new System.Drawing.Size(40, 16);
            this.label27.TabIndex = 51;
            this.label27.Text = "悬料";
            // 
            // c1NumericEdit25
            // 
            this.c1NumericEdit25.Culture = 127;
            this.c1NumericEdit25.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "Bengliao", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit25.EmptyAsNull = true;
            this.c1NumericEdit25.Font = new System.Drawing.Font("宋体", 10F);
            this.c1NumericEdit25.Location = new System.Drawing.Point(64, 69);
            this.c1NumericEdit25.Name = "c1NumericEdit25";
            this.c1NumericEdit25.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit25.TabIndex = 3;
            this.c1NumericEdit25.Tag = null;
            this.c1NumericEdit25.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit25.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit24
            // 
            this.c1NumericEdit24.Culture = 127;
            this.c1NumericEdit24.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "Zuoliao", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit24.EmptyAsNull = true;
            this.c1NumericEdit24.Font = new System.Drawing.Font("宋体", 10F);
            this.c1NumericEdit24.Location = new System.Drawing.Point(64, 41);
            this.c1NumericEdit24.Name = "c1NumericEdit24";
            this.c1NumericEdit24.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit24.TabIndex = 2;
            this.c1NumericEdit24.Tag = null;
            this.c1NumericEdit24.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit24.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit23
            // 
            this.c1NumericEdit23.Culture = 127;
            this.c1NumericEdit23.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "Xuanliao", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit23.EmptyAsNull = true;
            this.c1NumericEdit23.Font = new System.Drawing.Font("宋体", 10F);
            this.c1NumericEdit23.Location = new System.Drawing.Point(64, 12);
            this.c1NumericEdit23.Name = "c1NumericEdit23";
            this.c1NumericEdit23.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit23.TabIndex = 1;
            this.c1NumericEdit23.Tag = null;
            this.c1NumericEdit23.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit23.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.label18);
            this.groupBox4.Controls.Add(this.label19);
            this.groupBox4.Controls.Add(this.c1NumericEdit30);
            this.groupBox4.Controls.Add(this.c1NumericEdit29);
            this.groupBox4.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox4.ForeColor = System.Drawing.Color.Blue;
            this.groupBox4.Location = new System.Drawing.Point(31, 599);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(196, 76);
            this.groupBox4.TabIndex = 5;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "料批";
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label18.ForeColor = System.Drawing.Color.Blue;
            this.label18.Location = new System.Drawing.Point(9, 51);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(56, 16);
            this.label18.TabIndex = 50;
            this.label18.Text = "深料线";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label19.ForeColor = System.Drawing.Color.Blue;
            this.label19.Location = new System.Drawing.Point(9, 25);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(40, 16);
            this.label19.TabIndex = 49;
            this.label19.Text = "全部";
            // 
            // c1NumericEdit30
            // 
            this.c1NumericEdit30.Culture = 127;
            this.c1NumericEdit30.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "Shenliaopi", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit30.EmptyAsNull = true;
            this.c1NumericEdit30.Font = new System.Drawing.Font("宋体", 10F);
            this.c1NumericEdit30.Location = new System.Drawing.Point(114, 47);
            this.c1NumericEdit30.Name = "c1NumericEdit30";
            this.c1NumericEdit30.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit30.TabIndex = 2;
            this.c1NumericEdit30.Tag = null;
            this.c1NumericEdit30.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit30.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit29
            // 
            this.c1NumericEdit29.Culture = 127;
            this.c1NumericEdit29.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "Liaopi", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit29.EmptyAsNull = true;
            this.c1NumericEdit29.Font = new System.Drawing.Font("宋体", 10F);
            this.c1NumericEdit29.Location = new System.Drawing.Point(114, 18);
            this.c1NumericEdit29.Name = "c1NumericEdit29";
            this.c1NumericEdit29.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit29.TabIndex = 1;
            this.c1NumericEdit29.Tag = null;
            this.c1NumericEdit29.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit29.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // groupBox6
            // 
            this.groupBox6.Controls.Add(this.label30);
            this.groupBox6.Controls.Add(this.label31);
            this.groupBox6.Controls.Add(this.label32);
            this.groupBox6.Controls.Add(this.c1NumericEdit28);
            this.groupBox6.Controls.Add(this.c1NumericEdit27);
            this.groupBox6.Controls.Add(this.c1NumericEdit26);
            this.groupBox6.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox6.ForeColor = System.Drawing.Color.Blue;
            this.groupBox6.Location = new System.Drawing.Point(248, 416);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new System.Drawing.Size(140, 95);
            this.groupBox6.TabIndex = 8;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "风口";
            // 
            // label30
            // 
            this.label30.AutoSize = true;
            this.label30.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label30.ForeColor = System.Drawing.Color.Blue;
            this.label30.Location = new System.Drawing.Point(6, 74);
            this.label30.Name = "label30";
            this.label30.Size = new System.Drawing.Size(40, 16);
            this.label30.TabIndex = 56;
            this.label30.Text = "小套";
            // 
            // label31
            // 
            this.label31.AutoSize = true;
            this.label31.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label31.ForeColor = System.Drawing.Color.Blue;
            this.label31.Location = new System.Drawing.Point(6, 48);
            this.label31.Name = "label31";
            this.label31.Size = new System.Drawing.Size(40, 16);
            this.label31.TabIndex = 55;
            this.label31.Text = "中套";
            // 
            // label32
            // 
            this.label32.AutoSize = true;
            this.label32.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label32.ForeColor = System.Drawing.Color.Blue;
            this.label32.Location = new System.Drawing.Point(6, 22);
            this.label32.Name = "label32";
            this.label32.Size = new System.Drawing.Size(40, 16);
            this.label32.TabIndex = 54;
            this.label32.Text = "大套";
            // 
            // c1NumericEdit28
            // 
            this.c1NumericEdit28.Culture = 127;
            this.c1NumericEdit28.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "Fengkouxiaotao", true));
            this.c1NumericEdit28.EmptyAsNull = true;
            this.c1NumericEdit28.Font = new System.Drawing.Font("宋体", 10F);
            this.c1NumericEdit28.Location = new System.Drawing.Point(64, 66);
            this.c1NumericEdit28.Name = "c1NumericEdit28";
            this.c1NumericEdit28.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit28.TabIndex = 3;
            this.c1NumericEdit28.Tag = null;
            this.c1NumericEdit28.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit28.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit27
            // 
            this.c1NumericEdit27.Culture = 127;
            this.c1NumericEdit27.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "Fengkouzhongtao", true));
            this.c1NumericEdit27.EmptyAsNull = true;
            this.c1NumericEdit27.Font = new System.Drawing.Font("宋体", 10F);
            this.c1NumericEdit27.Location = new System.Drawing.Point(64, 41);
            this.c1NumericEdit27.Name = "c1NumericEdit27";
            this.c1NumericEdit27.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit27.TabIndex = 2;
            this.c1NumericEdit27.Tag = null;
            this.c1NumericEdit27.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit27.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit26
            // 
            this.c1NumericEdit26.Culture = 127;
            this.c1NumericEdit26.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "Fengkoudatao", true));
            this.c1NumericEdit26.EmptyAsNull = true;
            this.c1NumericEdit26.Font = new System.Drawing.Font("宋体", 10F);
            this.c1NumericEdit26.Location = new System.Drawing.Point(64, 16);
            this.c1NumericEdit26.Name = "c1NumericEdit26";
            this.c1NumericEdit26.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit26.TabIndex = 1;
            this.c1NumericEdit26.Tag = null;
            this.c1NumericEdit26.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit26.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // groupBox7
            // 
            this.groupBox7.Controls.Add(this.label26);
            this.groupBox7.Controls.Add(this.label25);
            this.groupBox7.Controls.Add(this.label24);
            this.groupBox7.Controls.Add(this.label23);
            this.groupBox7.Controls.Add(this.label22);
            this.groupBox7.Controls.Add(this.label20);
            this.groupBox7.Controls.Add(this.label21);
            this.groupBox7.Controls.Add(this.c1NumericEdit22);
            this.groupBox7.Controls.Add(this.c1NumericEdit21);
            this.groupBox7.Controls.Add(this.c1NumericEdit20);
            this.groupBox7.Controls.Add(this.c1NumericEdit19);
            this.groupBox7.Controls.Add(this.c1NumericEdit18);
            this.groupBox7.Controls.Add(this.c1NumericEdit17);
            this.groupBox7.Controls.Add(this.c1NumericEdit16);
            this.groupBox7.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox7.ForeColor = System.Drawing.Color.Blue;
            this.groupBox7.Location = new System.Drawing.Point(248, 111);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new System.Drawing.Size(140, 200);
            this.groupBox7.TabIndex = 6;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "炉渣";
            // 
            // label26
            // 
            this.label26.AutoSize = true;
            this.label26.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label26.ForeColor = System.Drawing.Color.Blue;
            this.label26.Location = new System.Drawing.Point(6, 169);
            this.label26.Name = "label26";
            this.label26.Size = new System.Drawing.Size(16, 16);
            this.label26.TabIndex = 55;
            this.label26.Text = "S";
            // 
            // label25
            // 
            this.label25.AutoSize = true;
            this.label25.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label25.ForeColor = System.Drawing.Color.Blue;
            this.label25.Location = new System.Drawing.Point(6, 144);
            this.label25.Name = "label25";
            this.label25.Size = new System.Drawing.Size(48, 16);
            this.label25.TabIndex = 54;
            this.label25.Text = "AL2O3";
            // 
            // label24
            // 
            this.label24.AutoSize = true;
            this.label24.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label24.ForeColor = System.Drawing.Color.Blue;
            this.label24.Location = new System.Drawing.Point(6, 119);
            this.label24.Name = "label24";
            this.label24.Size = new System.Drawing.Size(32, 16);
            this.label24.TabIndex = 53;
            this.label24.Text = "MnO";
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label23.ForeColor = System.Drawing.Color.Blue;
            this.label23.Location = new System.Drawing.Point(6, 94);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(32, 16);
            this.label23.TabIndex = 52;
            this.label23.Text = "MgO";
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label22.ForeColor = System.Drawing.Color.Blue;
            this.label22.Location = new System.Drawing.Point(6, 69);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(32, 16);
            this.label22.TabIndex = 51;
            this.label22.Text = "FeO";
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label20.ForeColor = System.Drawing.Color.Blue;
            this.label20.Location = new System.Drawing.Point(6, 44);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(32, 16);
            this.label20.TabIndex = 50;
            this.label20.Text = "CaO";
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label21.ForeColor = System.Drawing.Color.Blue;
            this.label21.Location = new System.Drawing.Point(6, 19);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(40, 16);
            this.label21.TabIndex = 49;
            this.label21.Text = "SiO2";
            // 
            // c1NumericEdit22
            // 
            this.c1NumericEdit22.Culture = 127;
            this.c1NumericEdit22.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "S", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit22.EmptyAsNull = true;
            this.c1NumericEdit22.Font = new System.Drawing.Font("宋体", 10F);
            this.c1NumericEdit22.Location = new System.Drawing.Point(64, 166);
            this.c1NumericEdit22.Name = "c1NumericEdit22";
            this.c1NumericEdit22.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit22.TabIndex = 7;
            this.c1NumericEdit22.Tag = null;
            this.c1NumericEdit22.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit22.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit21
            // 
            this.c1NumericEdit21.Culture = 127;
            this.c1NumericEdit21.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "Al2O3", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit21.EmptyAsNull = true;
            this.c1NumericEdit21.Font = new System.Drawing.Font("宋体", 10F);
            this.c1NumericEdit21.Location = new System.Drawing.Point(64, 141);
            this.c1NumericEdit21.Name = "c1NumericEdit21";
            this.c1NumericEdit21.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit21.TabIndex = 6;
            this.c1NumericEdit21.Tag = null;
            this.c1NumericEdit21.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit21.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit20
            // 
            this.c1NumericEdit20.Culture = 127;
            this.c1NumericEdit20.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "MnO", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit20.EmptyAsNull = true;
            this.c1NumericEdit20.Font = new System.Drawing.Font("宋体", 10F);
            this.c1NumericEdit20.Location = new System.Drawing.Point(64, 116);
            this.c1NumericEdit20.Name = "c1NumericEdit20";
            this.c1NumericEdit20.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit20.TabIndex = 5;
            this.c1NumericEdit20.Tag = null;
            this.c1NumericEdit20.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit20.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit19
            // 
            this.c1NumericEdit19.Culture = 127;
            this.c1NumericEdit19.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "MgO", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit19.EmptyAsNull = true;
            this.c1NumericEdit19.Font = new System.Drawing.Font("宋体", 10F);
            this.c1NumericEdit19.Location = new System.Drawing.Point(64, 91);
            this.c1NumericEdit19.Name = "c1NumericEdit19";
            this.c1NumericEdit19.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit19.TabIndex = 4;
            this.c1NumericEdit19.Tag = null;
            this.c1NumericEdit19.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit19.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit18
            // 
            this.c1NumericEdit18.Culture = 127;
            this.c1NumericEdit18.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "FeO", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit18.EmptyAsNull = true;
            this.c1NumericEdit18.Font = new System.Drawing.Font("宋体", 10F);
            this.c1NumericEdit18.Location = new System.Drawing.Point(64, 66);
            this.c1NumericEdit18.Name = "c1NumericEdit18";
            this.c1NumericEdit18.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit18.TabIndex = 3;
            this.c1NumericEdit18.Tag = null;
            this.c1NumericEdit18.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit18.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit17
            // 
            this.c1NumericEdit17.Culture = 127;
            this.c1NumericEdit17.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "CaO", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit17.EmptyAsNull = true;
            this.c1NumericEdit17.Font = new System.Drawing.Font("宋体", 10F);
            this.c1NumericEdit17.Location = new System.Drawing.Point(64, 41);
            this.c1NumericEdit17.Name = "c1NumericEdit17";
            this.c1NumericEdit17.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit17.TabIndex = 2;
            this.c1NumericEdit17.Tag = null;
            this.c1NumericEdit17.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit17.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit16
            // 
            this.c1NumericEdit16.Culture = 127;
            this.c1NumericEdit16.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "SiO2", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit16.EmptyAsNull = true;
            this.c1NumericEdit16.Font = new System.Drawing.Font("宋体", 10F);
            this.c1NumericEdit16.Location = new System.Drawing.Point(64, 16);
            this.c1NumericEdit16.Name = "c1NumericEdit16";
            this.c1NumericEdit16.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit16.TabIndex = 1;
            this.c1NumericEdit16.Tag = null;
            this.c1NumericEdit16.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit16.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label17);
            this.groupBox2.Controls.Add(this.label16);
            this.groupBox2.Controls.Add(this.c1NumericEdit15);
            this.groupBox2.Controls.Add(this.c1NumericEdit14);
            this.groupBox2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox2.ForeColor = System.Drawing.Color.Blue;
            this.groupBox2.Location = new System.Drawing.Point(31, 517);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(196, 76);
            this.groupBox2.TabIndex = 4;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "废铁";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label17.ForeColor = System.Drawing.Color.Blue;
            this.label17.Location = new System.Drawing.Point(9, 54);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(40, 16);
            this.label17.TabIndex = 48;
            this.label17.Text = "外购";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label16.ForeColor = System.Drawing.Color.Blue;
            this.label16.Location = new System.Drawing.Point(9, 28);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(40, 16);
            this.label16.TabIndex = 47;
            this.label16.Text = "全部";
            // 
            // c1NumericEdit15
            // 
            this.c1NumericEdit15.Culture = 127;
            this.c1NumericEdit15.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "Waigoufeitie", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit15.EmptyAsNull = true;
            this.c1NumericEdit15.Font = new System.Drawing.Font("宋体", 10F);
            this.c1NumericEdit15.Location = new System.Drawing.Point(114, 47);
            this.c1NumericEdit15.Name = "c1NumericEdit15";
            this.c1NumericEdit15.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit15.TabIndex = 2;
            this.c1NumericEdit15.Tag = null;
            this.c1NumericEdit15.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit15.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit14
            // 
            this.c1NumericEdit14.Culture = 127;
            this.c1NumericEdit14.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "Feitie", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit14.EmptyAsNull = true;
            this.c1NumericEdit14.Font = new System.Drawing.Font("宋体", 10F);
            this.c1NumericEdit14.Location = new System.Drawing.Point(114, 21);
            this.c1NumericEdit14.Name = "c1NumericEdit14";
            this.c1NumericEdit14.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit14.TabIndex = 1;
            this.c1NumericEdit14.Tag = null;
            this.c1NumericEdit14.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit14.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit13
            // 
            this.c1NumericEdit13.EmptyAsNull = true;
            this.c1NumericEdit13.Location = new System.Drawing.Point(114, 12);
            this.c1NumericEdit13.Name = "c1NumericEdit13";
            this.c1NumericEdit13.Size = new System.Drawing.Size(67, 21);
            this.c1NumericEdit13.TabIndex = 29;
            this.c1NumericEdit13.Tag = null;
            this.c1NumericEdit13.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            // 
            // c1DockingTab1
            // 
            this.c1DockingTab1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(128)))));
            this.c1DockingTab1.Controls.Add(this.休风);
            this.c1DockingTab1.Controls.Add(this.慢风);
            this.c1DockingTab1.Font = new System.Drawing.Font("宋体", 12F);
            this.c1DockingTab1.Location = new System.Drawing.Point(408, 114);
            this.c1DockingTab1.Name = "c1DockingTab1";
            this.c1DockingTab1.Size = new System.Drawing.Size(467, 555);
            this.c1DockingTab1.TabAreaBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.c1DockingTab1.TabIndex = 10;
            this.c1DockingTab1.TabsSpacing = 0;
            this.c1DockingTab1.VisualStyle = C1.Win.C1Command.VisualStyle.Custom;
            this.c1DockingTab1.VisualStyleBase = C1.Win.C1Command.VisualStyle.OfficeXP;
            // 
            // 休风
            // 
            this.休风.Controls.Add(this.c1FlexGrid1);
            this.休风.Location = new System.Drawing.Point(1, 30);
            this.休风.Name = "休风";
            this.休风.Size = new System.Drawing.Size(465, 524);
            this.休风.TabBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.休风.TabIndex = 0;
            this.休风.Text = "休风";
            // 
            // c1FlexGrid1
            // 
            this.c1FlexGrid1.AllowAddNew = true;
            this.c1FlexGrid1.AllowDelete = true;
            this.c1FlexGrid1.AllowDragging = C1.Win.C1FlexGrid.AllowDraggingEnum.None;
            this.c1FlexGrid1.AllowSorting = C1.Win.C1FlexGrid.AllowSortingEnum.None;
            this.c1FlexGrid1.AutoClipboard = true;
            this.c1FlexGrid1.AutoGenerateColumns = false;
            this.c1FlexGrid1.AutoResize = false;
            this.c1FlexGrid1.ColumnInfo = resources.GetString("c1FlexGrid1.ColumnInfo");
            this.c1FlexGrid1.DataSource = this.bindingSource1;
            this.c1FlexGrid1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.c1FlexGrid1.ExtendLastCol = true;
            this.c1FlexGrid1.Font = new System.Drawing.Font("宋体", 12F);
            this.c1FlexGrid1.KeyActionEnter = C1.Win.C1FlexGrid.KeyActionEnum.MoveAcross;
            this.c1FlexGrid1.Location = new System.Drawing.Point(0, 0);
            this.c1FlexGrid1.Name = "c1FlexGrid1";
            this.c1FlexGrid1.Rows.Count = 1;
            this.c1FlexGrid1.Rows.DefaultSize = 23;
            this.c1FlexGrid1.SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.Row;
            this.c1FlexGrid1.Size = new System.Drawing.Size(465, 524);
            this.c1FlexGrid1.StyleInfo = resources.GetString("c1FlexGrid1.StyleInfo");
            this.c1FlexGrid1.TabIndex = 31;
            this.c1FlexGrid1.Tag = "";
            this.c1FlexGrid1.AfterAddRow += new C1.Win.C1FlexGrid.RowColEventHandler(this.c1FlexGrid1_AfterAddRow);
            this.c1FlexGrid1.BeforeDeleteRow += new C1.Win.C1FlexGrid.RowColEventHandler(this.c1FlexGrid1_BeforeDeleteRow);
            // 
            // bindingSource1
            // 
            this.bindingSource1.DataMember = "休风";
            this.bindingSource1.DataSource = this.dataSetXh1;
            // 
            // dataSetXh1
            // 
            this.dataSetXh1.DataSetName = "DataSetXh1";
            this.dataSetXh1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // 慢风
            // 
            this.慢风.CaptionText = "慢风";
            this.慢风.Controls.Add(this.c1FlexGrid2);
            this.慢风.Location = new System.Drawing.Point(1, 30);
            this.慢风.Name = "慢风";
            this.慢风.Size = new System.Drawing.Size(465, 524);
            this.慢风.TabBackColor = System.Drawing.Color.Teal;
            this.慢风.TabBackColorSelected = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.慢风.TabIndex = 1;
            this.慢风.Text = "慢风";
            // 
            // c1FlexGrid2
            // 
            this.c1FlexGrid2.AllowAddNew = true;
            this.c1FlexGrid2.AllowDelete = true;
            this.c1FlexGrid2.AllowDragging = C1.Win.C1FlexGrid.AllowDraggingEnum.None;
            this.c1FlexGrid2.AllowSorting = C1.Win.C1FlexGrid.AllowSortingEnum.None;
            this.c1FlexGrid2.AutoClipboard = true;
            this.c1FlexGrid2.AutoResize = false;
            this.c1FlexGrid2.ColumnInfo = resources.GetString("c1FlexGrid2.ColumnInfo");
            this.c1FlexGrid2.DataSource = this.bindingSource2;
            this.c1FlexGrid2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.c1FlexGrid2.ExtendLastCol = true;
            this.c1FlexGrid2.Font = new System.Drawing.Font("宋体", 12F);
            this.c1FlexGrid2.KeyActionEnter = C1.Win.C1FlexGrid.KeyActionEnum.MoveAcross;
            this.c1FlexGrid2.Location = new System.Drawing.Point(0, 0);
            this.c1FlexGrid2.Name = "c1FlexGrid2";
            this.c1FlexGrid2.Rows.Count = 1;
            this.c1FlexGrid2.Rows.DefaultSize = 23;
            this.c1FlexGrid2.SelectionMode = C1.Win.C1FlexGrid.SelectionModeEnum.Row;
            this.c1FlexGrid2.Size = new System.Drawing.Size(465, 524);
            this.c1FlexGrid2.StyleInfo = resources.GetString("c1FlexGrid2.StyleInfo");
            this.c1FlexGrid2.TabIndex = 33;
            this.c1FlexGrid2.Tag = "";
            this.c1FlexGrid2.AfterAddRow += new C1.Win.C1FlexGrid.RowColEventHandler(this.c1FlexGrid1_AfterAddRow);
            this.c1FlexGrid2.BeforeDeleteRow += new C1.Win.C1FlexGrid.RowColEventHandler(this.c1FlexGrid1_BeforeDeleteRow);
            // 
            // bindingSource2
            // 
            this.bindingSource2.DataMember = "慢风";
            this.bindingSource2.DataSource = this.dataSetXh1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 21.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.ForeColor = System.Drawing.Color.Blue;
            this.label1.Location = new System.Drawing.Point(307, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(319, 29);
            this.label1.TabIndex = 15;
            this.label1.Text = "高 炉 日 消 耗 录 入";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.ForeColor = System.Drawing.Color.Blue;
            this.label2.Location = new System.Drawing.Point(279, 56);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(47, 19);
            this.label2.TabIndex = 16;
            this.label2.Text = "高炉";
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.label33);
            this.groupBox5.Controls.Add(this.label34);
            this.groupBox5.Controls.Add(this.label35);
            this.groupBox5.Controls.Add(this.c1NumericEdit31);
            this.groupBox5.Controls.Add(this.c1NumericEdit32);
            this.groupBox5.Controls.Add(this.c1NumericEdit33);
            this.groupBox5.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox5.ForeColor = System.Drawing.Color.Blue;
            this.groupBox5.Location = new System.Drawing.Point(248, 517);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(140, 152);
            this.groupBox5.TabIndex = 9;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "渣口";
            // 
            // label33
            // 
            this.label33.AutoSize = true;
            this.label33.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label33.ForeColor = System.Drawing.Color.Blue;
            this.label33.Location = new System.Drawing.Point(6, 75);
            this.label33.Name = "label33";
            this.label33.Size = new System.Drawing.Size(40, 16);
            this.label33.TabIndex = 59;
            this.label33.Text = "小套";
            // 
            // label34
            // 
            this.label34.AutoSize = true;
            this.label34.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label34.ForeColor = System.Drawing.Color.Blue;
            this.label34.Location = new System.Drawing.Point(6, 49);
            this.label34.Name = "label34";
            this.label34.Size = new System.Drawing.Size(40, 16);
            this.label34.TabIndex = 58;
            this.label34.Text = "中套";
            // 
            // label35
            // 
            this.label35.AutoSize = true;
            this.label35.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label35.ForeColor = System.Drawing.Color.Blue;
            this.label35.Location = new System.Drawing.Point(6, 23);
            this.label35.Name = "label35";
            this.label35.Size = new System.Drawing.Size(40, 16);
            this.label35.TabIndex = 57;
            this.label35.Text = "大套";
            // 
            // c1NumericEdit31
            // 
            this.c1NumericEdit31.Culture = 127;
            this.c1NumericEdit31.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "Zhakouxiaotao", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit31.EmptyAsNull = true;
            this.c1NumericEdit31.Font = new System.Drawing.Font("宋体", 10F);
            this.c1NumericEdit31.Location = new System.Drawing.Point(64, 66);
            this.c1NumericEdit31.Name = "c1NumericEdit31";
            this.c1NumericEdit31.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit31.TabIndex = 3;
            this.c1NumericEdit31.Tag = null;
            this.c1NumericEdit31.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit31.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit32
            // 
            this.c1NumericEdit32.Culture = 127;
            this.c1NumericEdit32.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "Zhakouzhongtao", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit32.EmptyAsNull = true;
            this.c1NumericEdit32.Font = new System.Drawing.Font("宋体", 10F);
            this.c1NumericEdit32.Location = new System.Drawing.Point(64, 41);
            this.c1NumericEdit32.Name = "c1NumericEdit32";
            this.c1NumericEdit32.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit32.TabIndex = 2;
            this.c1NumericEdit32.Tag = null;
            this.c1NumericEdit32.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit32.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // c1NumericEdit33
            // 
            this.c1NumericEdit33.Culture = 127;
            this.c1NumericEdit33.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "Zhakoudatao", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.c1NumericEdit33.EmptyAsNull = true;
            this.c1NumericEdit33.Font = new System.Drawing.Font("宋体", 10F);
            this.c1NumericEdit33.Location = new System.Drawing.Point(64, 16);
            this.c1NumericEdit33.Name = "c1NumericEdit33";
            this.c1NumericEdit33.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit33.TabIndex = 1;
            this.c1NumericEdit33.Tag = null;
            this.c1NumericEdit33.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            this.c1NumericEdit33.KeyDown += new System.Windows.Forms.KeyEventHandler(this.c1_KeyDown);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.button1.Location = new System.Drawing.Point(369, 52);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(112, 28);
            this.button1.TabIndex = 10;
            this.button1.Text = "提取数据";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // 慢风TableAdapter
            // 
            this.慢风TableAdapter.ClearBeforeFill = true;
            // 
            // 休风TableAdapter
            // 
            this.休风TableAdapter.ClearBeforeFill = true;
            // 
            // label36
            // 
            this.label36.AutoSize = true;
            this.label36.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label36.ForeColor = System.Drawing.Color.Blue;
            this.label36.Location = new System.Drawing.Point(5, 279);
            this.label36.Name = "label36";
            this.label36.Size = new System.Drawing.Size(72, 16);
            this.label36.TabIndex = 52;
            this.label36.Text = "焦粉水份";
            // 
            // c1NumericEdit7
            // 
            this.c1NumericEdit7.Culture = 127;
            this.c1NumericEdit7.DataBindings.Add(new System.Windows.Forms.Binding("Value", this.xiaohaoManagerBindingSource, "JiaofenShuiFen", true));
            this.c1NumericEdit7.EmptyAsNull = true;
            this.c1NumericEdit7.Location = new System.Drawing.Point(110, 275);
            this.c1NumericEdit7.Name = "c1NumericEdit7";
            this.c1NumericEdit7.Size = new System.Drawing.Size(67, 23);
            this.c1NumericEdit7.TabIndex = 51;
            this.c1NumericEdit7.Tag = null;
            this.c1NumericEdit7.VisibleButtons = C1.Win.C1Input.DropDownControlButtonFlags.None;
            // 
            // xiaohaoManagerBindingSource
            // 
            this.xiaohaoManagerBindingSource.DataSource = typeof(LTZN.XiaohaoGaolu.Xiaohao);
            // 
            // XiaohaoGaolu
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(7, 16);
            this.AutoScroll = true;
            this.BackColor = System.Drawing.Color.DarkSeaGreen;
            this.ClientSize = new System.Drawing.Size(887, 702);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.c1DockingTab1);
            this.Controls.Add(this.groupBox7);
            this.Controls.Add(this.groupBox6);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.dateTimePicker1);
            this.Font = new System.Drawing.Font("宋体", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Name = "XiaohaoGaolu";
            this.Text = "高炉日消耗录入";
            this.Load += new System.EventHandler(this.XiaohaoGaolu_Load);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.XiaohaoGaolu_FormClosed);
            this.Leave += new System.EventHandler(this.XiaohaoGaolu_Leave);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit34)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit12)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit11)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit10)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit9)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit8)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit6)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit5)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit4)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit3)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit0)).EndInit();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit25)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit24)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit23)).EndInit();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit30)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit29)).EndInit();
            this.groupBox6.ResumeLayout(false);
            this.groupBox6.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit28)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit27)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit26)).EndInit();
            this.groupBox7.ResumeLayout(false);
            this.groupBox7.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit22)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit21)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit20)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit19)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit18)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit17)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit16)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit15)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit14)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit13)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1DockingTab1)).EndInit();
            this.c1DockingTab1.ResumeLayout(false);
            this.休风.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.c1FlexGrid1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataSetXh1)).EndInit();
            this.慢风.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.c1FlexGrid2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource2)).EndInit();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit31)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit32)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit33)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.c1NumericEdit7)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.xiaohaoManagerBindingSource)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void XiaohaoGaolu_Load(object sender, EventArgs e)
        {
            this.慢风TableAdapter.FillByPk(this.dataSetXh1.慢风, DateTime.Today, 1);
            this.休风TableAdapter.FillByPk(this.dataSetXh1.休风, DateTime.Today, 1);
            XiaohaoManager xhm=new XiaohaoManager();
            xhm.Add(DateTime.Today, 1);
            this.xiaohaoManagerBindingSource.DataSource = xhm;


            string xffl = "";
            foreach (string s in Properties.Settings.Default.休风分类)
            {
                xffl += s + "|";
            }
            this.c1FlexGrid1.Cols[3].ComboList = xffl;

            string mffl = "";
            foreach (string s in Properties.Settings.Default.慢风分类)
            {
                mffl += s + "|";
            }
            this.c1FlexGrid2.Cols[3].ComboList = mffl;
            
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            getData(this.dateTimePicker1.Value.Date, Convert.ToInt32(comboBox1.Text));  
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            getData(this.dateTimePicker1.Value.Date, Convert.ToInt32(comboBox1.Text));    
        }

        private void getData(DateTime riqi,Int32 gaolu)
        { 
            XiaohaoManager xhm = this.xiaohaoManagerBindingSource.DataSource as XiaohaoManager;
            if (canUpdate)
            {
                this.c1FlexGrid1.FinishEditing();
                this.c1FlexGrid1.FinishEditing();

                this.bindingSource1.CurrencyManager.EndCurrentEdit();
                this.bindingSource2.CurrencyManager.EndCurrentEdit();

                this.慢风TableAdapter.Update(this.dataSetXh1.慢风);
                this.休风TableAdapter.Update(this.dataSetXh1.休风);
                xhm.save();
            }
           
            xhm.Add(riqi, gaolu);
            this.xiaohaoManagerBindingSource.ResetCurrentItem();
            this.慢风TableAdapter.FillByPk(this.dataSetXh1.慢风, riqi, gaolu);
            this.休风TableAdapter.FillByPk(this.dataSetXh1.休风, riqi, gaolu);
        }

        private void c1FlexGrid1_AfterAddRow(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            C1FlexGrid  cf=(C1FlexGrid)sender;
            cf[e.Row,0] = this.dateTimePicker1.Value.Date;
            cf[e.Row, 1] = Convert.ToInt32(comboBox1.Text);
        }


        private void c1FlexGrid1_Validated(object sender, EventArgs e)
        {
            C1FlexGrid cf = (C1FlexGrid)sender;
            if (canUpdate)
            {
                if (sender == null)
                    return;
                cf.FinishEditing();
                CurrencyManager cm = (CurrencyManager)BindingContext[cf.DataSource, cf.DataMember];
                cm.EndCurrentEdit();
                this.休风TableAdapter.Update(this.dataSetXh1.休风);
            }
        }

        private void c1FlexGrid2_Validated(object sender, EventArgs e)
        {
            C1FlexGrid cf = (C1FlexGrid)sender;
            if (canUpdate)
            {
                if (sender == null)
                    return;
                cf.FinishEditing();
                CurrencyManager cm = (CurrencyManager)BindingContext[cf.DataSource, cf.DataMember];
                cm.EndCurrentEdit();
                this.慢风TableAdapter.Update(this.dataSetXh1.慢风);
            }
        }

        private void XiaohaoGaolu_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (canUpdate)
            {
                this.c1FlexGrid1.FinishEditing();
                this.c1FlexGrid1.FinishEditing();

                this.bindingSource1.CurrencyManager.EndCurrentEdit();
                this.bindingSource2.CurrencyManager.EndCurrentEdit();

                this.慢风TableAdapter.Update(this.dataSetXh1.慢风);
                this.休风TableAdapter.Update(this.dataSetXh1.休风);

                XiaohaoManager xhm = this.xiaohaoManagerBindingSource.DataSource as XiaohaoManager;
                if (xhm!=null)
                   xhm.save();
            }
        }

        //提取数据
        private void button1_Click(object sender, EventArgs e)
        {
            OracleConnection cn = new OracleConnection();
            cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
            cn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = cn;
            cmd.CommandText = "SELECT TRUNC(AVG(ZHASIO2),2),TRUNC(AVG(ZHACAO),2),TRUNC(AVG(ZHAMGO),2),TRUNC(AVG(ZHAAL2O3),2),TRUNC(AVG(ZHAS),2),TRUNC(AVG(ZHATIO2),3) FROM DDLUCI WHERE TRUNC(ZDSJ)>=:RQ AND GAOLU=:GAOLU AND DKSJ IS NOT NULL";
            cmd.Parameters.Add(":RQ", OracleType.DateTime).Value = dateTimePicker1.Value.Date;
            cmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = Convert.ToInt32(comboBox1.Text);
            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                c1NumericEdit16.Value = dr.IsDBNull(0) ? DBNull.Value : (object)dr.GetDecimal(0);
                c1NumericEdit17.Value = dr.IsDBNull(1) ? DBNull.Value : (object)dr.GetDecimal(1);
                c1NumericEdit19.Value = dr.IsDBNull(2) ? DBNull.Value : (object)dr.GetDecimal(2);
                c1NumericEdit21.Value = dr.IsDBNull(3) ? DBNull.Value : (object)dr.GetDecimal(3);
                c1NumericEdit22.Value = dr.IsDBNull(4) ? DBNull.Value : (object)dr.GetDecimal(4);

            }
            dr.Close();
            cmd.CommandText = "SELECT  MC,SUM(NVL(BAIBAN,0)+NVL(ZHONGBAN,0)+NVL(YEBAN,0)) FROM RBXIAOHAO WHERE TRUNC(SJ)=:RQ AND GAOLU=:GAOLU GROUP BY MC";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                string mc = "";
                if (!dr.IsDBNull(0))
                    mc = dr.GetString(0);
                switch (mc)
                {                    
                    case "富氧量":
                        c1NumericEdit0.Value = dr.IsDBNull(1) ? DBNull.Value : (object)dr.GetDecimal(1);
                        break;
                    case "焦丁":
                        c1NumericEdit1.Value = dr.IsDBNull(1) ? DBNull.Value : (object)dr.GetDecimal(1);
                        break;
                }
            }
            dr.Close();
            cmd.CommandText = "select trunc(avg(lfll),0),trunc(avg(rfwd),0),trunc(avg(rfyl),0),trunc(avg(ldwd),0),trunc(avg(ldyl),0),trunc(avg(fyll),0) from rbcaozuo where trunc(sj)=:RQ and gaolu=:GAOLU";
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                c1NumericEdit2.Value = dr.IsDBNull(0) ? DBNull.Value : (object)(dr.GetDouble(0)*60);
                c1NumericEdit3.Value = dr.IsDBNull(1) ? DBNull.Value : (object)dr.GetDouble(1);
                c1NumericEdit4.Value = dr.IsDBNull(2) ? DBNull.Value : (object)dr.GetDouble(2);
                c1NumericEdit5.Value = dr.IsDBNull(3) ? DBNull.Value : (object)dr.GetDouble(3);
                c1NumericEdit6.Value = dr.IsDBNull(4) ? DBNull.Value : (object)dr.GetDouble(4);

                /*
                if (冷风流量 > 0 || 富氧量 > 0)
                {
                    富氧率 = (富氧量 / (富氧量 + 冷风流量)) * 0.79 * 100 / 60;  //计算的富氧率
                    富氧率 = double.Parse(富氧率.Value.ToString("0.00"));
                }
                else
                    富氧率 = 0;
                 */
            }
            dr.Close();

            cmd.CommandText = "select count(*) from ddluci where trunc(zdsj)=:RQ and gaolu=:GAOLU and dksj is not null";
            try { c1NumericEdit9.Value = cmd.ExecuteScalar(); } //炉数
            catch { c1NumericEdit9.Value =DBNull.Value;}
           
            cmd.CommandText = "select trunc(avg(CO2),2),trunc(avg(CO),2) from ddmq where trunc(sj)=:RQ and gaolu=:GAOLU";
            dr=cmd.ExecuteReader();
            if (dr.Read())
            {
                c1NumericEdit10.Value = dr.IsDBNull(0) ? DBNull.Value : (object)dr.GetDouble(0);
                c1NumericEdit11.Value = dr.IsDBNull(1) ? DBNull.Value : (object)dr.GetDouble(1);
            }
            dr.Close();

            cmd.CommandText = "select count(*) from rbliaoxian where trunc(t)=:RQ and gaolu=:GAOLU";
            try { c1NumericEdit29.Value = cmd.ExecuteScalar(); }//全部料线
            catch { c1NumericEdit29.Value = DBNull.Value; }
            cmd.CommandText = "select count(*) from rbliaoxian where trunc(t)=:RQ and gaolu=:GAOLU　and liaoqian>1900";
            try { c1NumericEdit30.Value = cmd.ExecuteScalar(); } //深料线
            catch { c1NumericEdit30.Value = DBNull.Value; }

            cmd.CommandText = "select count(*) from rbcaozuo where  TRUNC(SJ)=:RQ AND GAOLU =:GAOLU 　and ldyl>=30";
            try { c1NumericEdit12.Value = cmd.ExecuteScalar(); } //深料线
            catch { c1NumericEdit12.Value = DBNull.Value; }
            
            cn.Close();
           
            
        }

        private void c1_KeyDown(object sender, KeyEventArgs e)
        {
            C1NumericEdit cn=(C1NumericEdit)sender;
            switch (e.KeyCode)
            {
                case Keys.Delete:
                    cn.Value = System.DBNull.Value;
                    break;
                case Keys.Enter:
                    cn.Parent.Parent.SelectNextControl(cn, true, true, true, false);
                    break;
                case Keys.Down:
                    cn.Parent.Parent.SelectNextControl(cn, true, true, true, false);
                    break;
                case Keys.Up:
                    cn.Parent.Parent.SelectNextControl(cn, false, true, true, false);
                    break;
            }
        }

        private void XiaohaoGaolu_Leave(object sender, EventArgs e)
        {
            if (canUpdate)
            {
                this.c1FlexGrid1.FinishEditing();
                this.c1FlexGrid1.FinishEditing();

                this.bindingSource1.CurrencyManager.EndCurrentEdit();
                this.bindingSource2.CurrencyManager.EndCurrentEdit();

                this.慢风TableAdapter.Update(this.dataSetXh1.慢风);
                this.休风TableAdapter.Update(this.dataSetXh1.休风);
                XiaohaoManager xhm = this.xiaohaoManagerBindingSource.DataSource as XiaohaoManager;
                if (xhm != null)
                    xhm.save();
            }

        }

        private void c1FlexGrid1_BeforeDeleteRow(object sender, RowColEventArgs e)
        {
            if (MessageBox.Show("您确实要删除记录吗？", "询问", MessageBoxButtons.YesNo) == DialogResult.No)
                e.Cancel = true;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
		
	}
}
