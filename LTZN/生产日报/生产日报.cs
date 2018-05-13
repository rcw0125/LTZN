using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;
using C1.Win.C1FlexGrid;
using System.Diagnostics;
using ZHCDB;

namespace LTZN.生产日报
{
    public partial class 生产日报 : Form
    {
        public List<int> AuthGaolu = new List<int>();
      //  private int gaolu = 0; //高炉
        private bool xianzhi = true; //限制

        private RbxfTable rbxfTable = new RbxfTable(); //休放风情况

        public 生产日报()
        {
            InitializeComponent();
            this.生产数据1.RBXIAOHAO.ColumnChanged += new DataColumnChangeEventHandler(RBXIAOHAO_ColumnChanged);
            操作参数平均.SelectedIndex = 1;
            LtznUserManager.instance.RegisterHandler(this, instance_UserChanged);
        }
        void instance_UserChanged(LtznUser ltznUser)
        {
            AuthGaolu.Clear();
            xianzhi = true;
            if (ltznUser != null)
            {
                if (ltznUser.IsInRole("1高炉"))
                {
                    AuthGaolu.Add(1);           
                }
                if (ltznUser.IsInRole("2高炉"))
                {
                    AuthGaolu.Add(2);
                }
                if (ltznUser.IsInRole("3高炉"))
                {
                    AuthGaolu.Add(3);
                }
                if (ltznUser.IsInRole("4高炉"))
                {
                    AuthGaolu.Add(4);
                }
                if (ltznUser.IsInRole("5高炉"))
                {
                    AuthGaolu.Add(5);
                }
                if (ltznUser.IsInRole("6高炉"))
                {
                    AuthGaolu.Add(6);
                }
                if (ltznUser.IsInRole("无时间限制"))
                {
                    xianzhi = false;
                }
                if (AuthGaolu.Count > 0)
                    this.comboBox1.Text = AuthGaolu[0].ToString();
                else
                    this.comboBox1.Text = "1";

            }
            quanxian();
        }

        void RBXIAOHAO_ColumnChanged(object sender, DataColumnChangeEventArgs e)
        {
            if (e.Column.ColumnName == "BAIBAN" || e.Column.ColumnName == "ZHONGBAN" || e.Column.ColumnName == "YEBAN")
            {
                decimal yeban = 0, baiban = 0, zhongban = 0;
                if (e.Row["YEBAN"] != DBNull.Value) yeban = (decimal)e.Row["YEBAN"];
                if (e.Row["BAIBAN"] != DBNull.Value) baiban = (decimal)e.Row["BAIBAN"];
                if (e.Row["ZHONGBAN"] != DBNull.Value) zhongban = (decimal)e.Row["ZHONGBAN"];
                e.Row["LEIJI"] = yeban + baiban + zhongban;
            }
        }

        //保存数据
        private void button1_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;

            DateTime RQ = this.dateTimePicker1.Value.Date;
            int gaolu = Convert.ToInt32(this.comboBox1.Text);

            //料线查询
            switch (操作参数平均.SelectedTab.Name)
            {
                case "SLQK":    //上料情况
                    if (this.生产数据1.RBLIAOXIANYB.DataSet.HasChanges() == true)
                    {
                        this.rbliaoxianybTableAdapter1.Update(生产数据1.RBLIAOXIANYB);
                    }
                    this.rbliaoxianybTableAdapter1.Update(生产数据1.RBLIAOXIANYB);
                    if (this.生产数据1.RBLIAOXIANZB.DataSet.HasChanges() == true)
                    {
                      
                        this.rbliaoxianzbTableAdapter1.Update(生产数据1.RBLIAOXIANZB);
                    }
                    if (this.生产数据1.RBLIAOXIANBB.DataSet.HasChanges() == true)
                    {

                        this.rbliaoxianbbTableAdapter1.Update(生产数据1.RBLIAOXIANBB);
                    }
                    break;
                case "TZQK":    //铁渣情况
                    this.ddluciTableAdapter1.Update(生产数据1.DDLUCI);
                    break;
                case "BDQK":    //变动情况
                    this.c1FlexGrid_biandong.FinishEditing();
                    CurrencyManager cmBiandong = (CurrencyManager)BindingContext[this.c1FlexGrid_biandong.DataSource, this.c1FlexGrid_biandong.DataMember];
                    cmBiandong.EndCurrentEdit();
                    foreach (生产数据.RBBDRow row in this.生产数据1.RBBD.Rows)
                    {
                        if (row.RowState == DataRowState.Added || row.RowState == DataRowState.Modified)
                        {
                            row.GAOLU = gaolu;
                            if (row.IsRQNull())
                                row.SJ = new DateTime(RQ.Year, RQ.Month, RQ.Day, DateTime.Now.Hour, DateTime.Now.Minute, 0);
                            else
                                row.SJ = new DateTime(RQ.Year, RQ.Month, RQ.Day, row.RQ.Hour, row.RQ.Minute, 0);
                            row.RQ = row.SJ;
                        }
                    }
                    this.rbbdTableAdapter1.Update(生产数据1.RBBD);
                    break;
                case "CZCS":    //操作参数
                    if (this.生产数据1.RBCAOZUO.DataSet.HasChanges() == true)
                    {
                        foreach (生产数据.RBCAOZUORow row in this.生产数据1.RBCAOZUO.Rows)
                        {
                            if (row.RowState == DataRowState.Added || row.RowState == DataRowState.Modified)
                            {
                                int t = 0;
                                try
                                {
                                    t = Convert.ToInt32(row["HOUR"]);
                                }
                                catch { }
                                row.SJ = new DateTime(RQ.Year, RQ.Month, RQ.Day, t, 0, 0);
                                row.GAOLU = gaolu;
                            }
                        }
                    this.c1FlexGrid_caozuo.Subtotal(AggregateEnum.Clear);
                   
                        this.rbcaozuoTableAdapter1.Update(生产数据1.RBCAOZUO);
                    }
                    break;
                case "ZLZD":    //装料制度
                        try
                        {
                            C1FlexGrid flex2 = this.c1FlexGrid_zhuangliao;
                            CurrencyManager cm2 = (CurrencyManager)BindingContext[flex2.DataSource, flex2.DataMember];
                            cm2.EndCurrentEdit();
                            flex2.FinishEditing();
              
                  if (this.生产数据1.RBZLBD.DataSet.HasChanges() == true)
                    {
                        foreach (生产数据.RBZLBDRow row in this.生产数据1.RBZLBD.Rows)
                        {
                            if (row.RowState == DataRowState.Added || row.RowState == DataRowState.Modified)
                            {
                                row.GAOLU = gaolu;
                                row.SJ = new DateTime(RQ.Year, RQ.Month, RQ.Day, DateTime.Now.Hour, DateTime.Now.Minute, 0);
                                if (!row.IsBLSJNull())
                                {
                                    row.BLSJ = new DateTime(RQ.Year, RQ.Month, RQ.Day, row.BLSJ.Hour, row.BLSJ.Minute, 0);
                                    if (row.BC == "中班" && row.BLSJ - row.SJ > TimeSpan.FromHours(8))
                                    {
                                        row.BLSJ.AddDays(-1);
                                    }
                                }
                                else
                                {
                                    row.BLSJ = row.SJ;
                                }
                            }
                        }
                        this.rbzlbdTableAdapter1.Update(生产数据1.RBZLBD);
                   }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.ToString());
                        }
                    break;
                case "FKQK":    //风口情况
                    {
                        C1FlexGrid flex = this.c1FlexGrid_fengkou;
                        flex.FinishEditing();
                        CurrencyManager cm1 = (CurrencyManager)BindingContext[flex.DataSource, flex.DataMember];
                        cm1.EndCurrentEdit();
                    }
                    foreach (生产数据.RBFKRow row in this.生产数据1.RBFK.Rows)
                    {
                        if (row.RowState == DataRowState.Added || row.RowState == DataRowState.Modified)
                        {
                            row.GAOLU = gaolu;
                            if (row.IsRQNull())
                                row.SJ = new DateTime(RQ.Year, RQ.Month, RQ.Day, DateTime.Now.Hour, DateTime.Now.Minute, 0);
                            else
                                row.SJ = new DateTime(RQ.Year, RQ.Month, RQ.Day, row.RQ.Hour, row.RQ.Minute, 0);
                            row.RQ = row.SJ;
                        }
                    }
                    this.rbfkTableAdapter1.Update(生产数据1.RBFK);
                    break;
                case "LQQK":    //冷却情况
                    //foreach (生产数据.RBLQRow row in this.生产数据1.RBLQ.Rows)
                    //{
                    //    if (row.RowState == DataRowState.Added || row.RowState == DataRowState.Modified)
                    //    {
                    //        row.GAOLU = gaolu;
                    //        row.SJ = this.dateTimePicker1.Value;
                    //    }
                    //}
                    this.c1FlexGrid_lengque.Subtotal(AggregateEnum.Clear);
                    this.rblqTableAdapter1.Update(生产数据1.RBLQ);
                    break;
                case "XHQK":    //消耗情况
                    foreach (生产数据.RBXIAOHAORow row in this.生产数据1.RBXIAOHAO.Rows)
                    {
                        if (row.RowState == DataRowState.Added || row.RowState == DataRowState.Modified)
                        {
                            row.SJ = RQ;
                            row.GAOLU = gaolu;
                        }
                    }
                    this.rbxiaohaoTableAdapter1.Update(生产数据1.RBXIAOHAO);
                    break;
                case "XFFF":    //休风放风
                    rbxfTable.Save();
                    //foreach (生产数据.RBXFRow row in this.生产数据1.RBXF.Rows)
                    //{
                    //    if (row.RowState == DataRowState.Added || row.RowState == DataRowState.Modified)
                    //    {
                    //        row.GAOLU = gaolu;
                    //        row.SJ = RQ;
                    //        if (!row.IsKSNull())
                    //            row.KS = new DateTime(RQ.Year, RQ.Month, RQ.Day, row.KS.Hour, row.KS.Minute, 0);
                    //        if (!row.IsZZNull())
                    //            row.ZZ = new DateTime(RQ.Year, RQ.Month, RQ.Day, row.ZZ.Hour, row.ZZ.Minute, 0);
                    //    }
                    //}
                    //this.rbxfTableAdapter1.Update(生产数据1.RBXF);
                    break;
                case "GZRZ":

                    updateBanRiZhi(Convert.ToInt32(this.comboBox1.Text), this.dateTimePicker1.Value.Date, "夜班", this.ybTxt.Text);

                    updateBanRiZhi(Convert.ToInt32(this.comboBox1.Text), this.dateTimePicker1.Value.Date, "白班", this.bbTxt.Text);

                    updateBanRiZhi(Convert.ToInt32(this.comboBox1.Text), this.dateTimePicker1.Value.Date, "中班", this.zbTxt.Text);


                    break;
                case "JSB":                   
                    OracleConnection cn1 = new OracleConnection();
                    cn1.ConnectionString = Properties.Settings.Default.ltznConnectionString;
                    cn1.Open();
                    OracleCommand selCmd = new OracleCommand();
                    selCmd.Connection = cn1;
                    selCmd.CommandType = CommandType.Text;
                    selCmd.CommandText = "SELECT Count(*) FROM RBJS WHERE TRUNC(SJ)=:RQ and gaolu=:GAOLU";
                    selCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = this.dateTimePicker1.Value.Date;
                    selCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = Convert.ToInt32(this.comboBox1.Text);
                    OracleNumber count = (OracleNumber)selCmd.ExecuteOracleScalar();
                    if (count.Value > 0)
                    {
                        OracleCommand updateCmd = new OracleCommand();
                        updateCmd.Connection = cn1;
                        updateCmd.CommandText = "UPDATE RBJS SET JS=:JS WHERE TRUNC(SJ)=:RQ and gaolu=:GAOLU";
                        updateCmd.Parameters.Add(":JS", OracleType.VarChar).Value = Convert.ToString(this.jsbTxt.Text);
                        updateCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = this.dateTimePicker1.Value.Date;
                        updateCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = Convert.ToInt32(this.comboBox1.Text);
                        updateCmd.ExecuteNonQuery();
                    }
                    else
                    {
                        OracleCommand insertCmd = new OracleCommand();
                        insertCmd.Connection = cn1;
                        insertCmd.CommandText = "INSERT INTO RBJS(SJ, GAOLU, JS) VALUES (:SJ, :GAOLU, :JS)";
                        insertCmd.Parameters.Add(":SJ", OracleType.DateTime).Value = this.dateTimePicker1.Value.Date;
                        insertCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = Convert.ToInt32(this.comboBox1.Text);
                        insertCmd.Parameters.Add(":JS", OracleType.VarChar).Value = Convert.ToString(this.jsbTxt.Text);
                        insertCmd.ExecuteNonQuery();
                    }

                    cn1.Close();
                    break;

            }
         //   button1.Enabled = false;
            this.Cursor = Cursors.Default;

        }

        //更新值班日志
        private void updateBanRiZhi(int gaolu, DateTime rq, string banci, string neirong)
        {
            OracleConnection cn = new OracleConnection();
            cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
            cn.Open();
            OracleCommand selCmd = new OracleCommand();
            selCmd.Connection = cn;
            selCmd.CommandText = "Select Count(*) from RBGZ WHERE (GAOLU = :GAOLU) AND (TRUNC(SJ) = :SJ) AND (BC = :BC)";
            selCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = gaolu;
            selCmd.Parameters.Add(":SJ", OracleType.DateTime).Value = rq;
            selCmd.Parameters.Add(":BC", OracleType.VarChar).Value = banci;
            OracleNumber count = (OracleNumber)selCmd.ExecuteOracleScalar();
            if (count.Value > 0)
            {
                OracleCommand updateCmd = new OracleCommand();
                updateCmd.Connection = cn;
                updateCmd.CommandText = "UPDATE RBGZ SET JL=:JL WHERE (GAOLU = :GAOLU) AND (TRUNC(SJ) = :SJ) AND (BC = :BC)";
                updateCmd.Parameters.Add(":JL", OracleType.VarChar).Value = neirong;
                updateCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = gaolu;
                updateCmd.Parameters.Add(":SJ", OracleType.DateTime).Value = rq;
                updateCmd.Parameters.Add(":BC", OracleType.VarChar).Value = banci;
                updateCmd.ExecuteNonQuery();
            }
            else
            {
                OracleCommand insertCmd = new OracleCommand();
                insertCmd.Connection = cn;
                insertCmd.CommandText = "INSERT INTO RBGZ(GAOLU,SJ,BC,JL)VALUES (:GAOLU,:SJ, :BC, :JL)";
                insertCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = gaolu;
                insertCmd.Parameters.Add(":SJ", OracleType.DateTime).Value = rq;
                insertCmd.Parameters.Add(":BC", OracleType.VarChar).Value = banci;
                insertCmd.Parameters.Add(":JL", OracleType.VarChar).Value = neirong;
                insertCmd.ExecuteNonQuery();
            }
            cn.Close();
        }

        private void quanxian()
        {
           

            if ((!xianzhi || this.dateTimePicker1.Value.Date > LtznUserManager.instance.Now - TimeSpan.FromHours(48)) && (AuthGaolu.Count > 0) && AuthGaolu.Contains(Convert.ToInt32(this.comboBox1.Text)))
            {
                btnSave.Visible = true;
                //button1.Enabled = true;

                button2.Visible = true;
                button2.Enabled = true;

                //上料情况
                c1FlexGrid_yeban.AllowAddNew = true;
                //  c1FlexGrid_yeban.AllowDelete = true;

                c1FlexGrid_yeban.AllowEditing = true;

                c1FlexGrid_baiban.AllowAddNew = true;
                // c1FlexGrid_baiban.AllowDelete = true;
                c1FlexGrid_baiban.AllowEditing = true;

                c1FlexGrid_zhongban.AllowAddNew = true;
                // c1FlexGrid_zhongban.AllowDelete = true;
                c1FlexGrid_zhongban.AllowEditing = true;

                //铁渣情况

                c1FlexGrid_tiezha.AllowEditing = true;

                //变动情况
                c1FlexGrid_biandong.AllowAddNew = true;
                // c1FlexGrid_biandong.AllowDelete = true;
                c1FlexGrid_biandong.AllowEditing = true;
                zlbdMenu.Enabled = true;

                //操作情况
                //c1FlexGrid_caozuo.AllowAddNew = true;
                // c1FlexGrid_caozuo.AllowDelete = true;
                c1FlexGrid_caozuo.AllowEditing = true;


                //装料制度
                c1FlexGrid_zhuangliao.AllowAddNew = true;
                c1FlexGrid_zhuangliao.AllowEditing = true;
                //   c1FlexGrid_zhuangliao.AllowDelete = true;
                zlzdMenu.Enabled = true;
                //风口情况
                c1FlexGrid_fengkou.AllowEditing = true;
                c1FlexGrid_fengkou.AllowAddNew = true;
                //  c1FlexGrid_fengkou.AllowDelete = true;
                fkMenu.Enabled = true;
                //冷却情况
                c1FlexGrid_lengque.AllowEditing = true;
                c1FlexGrid_lengque.AllowAddNew = true;
                //    c1FlexGrid_lengque.AllowDelete = true;
                lqMenu.Enabled = true;
                //消耗情况
                c1FlexGrid_xiaohao.AllowAddNew = true;
                // c1FlexGrid_xiaohao.AllowDelete = true;
                c1FlexGrid_xiaohao.AllowEditing = true;
                xiaoHaoMenu.Enabled = true;
                List<string> lists = new List<string>();
                高炉燃料比综合分析.LegendEnviroment.loadGLYL("高炉原料",out lists);
                StringBuilder sb=new StringBuilder();
                foreach(String s in lists)
                {
                   sb.Append(s+"|");
                }
                c1FlexGrid_xiaohao.Cols[2].ComboList = sb.ToString();
                //休风情况
                c1FlexGrid_xiufeng.AllowAddNew = true;
                //    c1FlexGrid_xiufeng.AllowDelete = true;
                c1FlexGrid_xiufeng.AllowEditing = true;
                xfMenu.Enabled = true;

                //工作记录
                ybTxt.ReadOnly = false;
                ybTxt.BackColor = Color.White;
                bbTxt.ReadOnly = false;
                bbTxt.BackColor = Color.White;
                zbTxt.ReadOnly = false;
                zbTxt.BackColor = Color.White;
                //记事本
                jsbTxt.ReadOnly = false;
                jsbTxt.BackColor = Color.White;
            }
            else
            {
                btnSave.Visible = false;
                // button1.Enabled = false;
                button2.Visible = false;
                button2.Enabled = false;


                //上料情况
                c1FlexGrid_yeban.AllowAddNew = false;
                c1FlexGrid_yeban.AllowDelete = false;
                c1FlexGrid_yeban.AllowEditing = false;

                c1FlexGrid_baiban.AllowAddNew = false;
                c1FlexGrid_baiban.AllowDelete = false;
                c1FlexGrid_baiban.AllowEditing = false;

                c1FlexGrid_zhongban.AllowAddNew = false;
                c1FlexGrid_zhongban.AllowDelete = false;
                c1FlexGrid_zhongban.AllowEditing = false;

                //铁渣情况

                c1FlexGrid_tiezha.AllowEditing = false;

                //变动情况
                c1FlexGrid_biandong.AllowAddNew = false;
                c1FlexGrid_biandong.AllowDelete = false;
                c1FlexGrid_biandong.AllowEditing = false;
                zlbdMenu.Enabled = false;

                //操作情况
                // c1FlexGrid_caozuo.AllowAddNew = false;
                //c1FlexGrid_caozuo.AllowDelete = false;
                c1FlexGrid_caozuo.AllowEditing = false;

                //装料制度
                //c1FlexGrid_zhuangliao.AllowAddNew = false;
                //c1FlexGrid_zhuangliao.AllowDelete = false;
                //c1FlexGrid_zhuangliao.AllowEditing = false;
                //zlzdMenu.Enabled = false;

                //风口情况
                c1FlexGrid_fengkou.AllowEditing = false;
                c1FlexGrid_fengkou.AllowAddNew = false;
                c1FlexGrid_fengkou.AllowDelete = false;
                fkMenu.Enabled = false;
                //冷却
                c1FlexGrid_lengque.AllowAddNew = false;
                c1FlexGrid_lengque.AllowDelete = false;
                c1FlexGrid_lengque.AllowEditing = false;
                lqMenu.Enabled = false;

                //消耗情况
                c1FlexGrid_xiaohao.AllowAddNew = false;
                c1FlexGrid_xiaohao.AllowDelete = false;
                c1FlexGrid_xiaohao.AllowEditing = false;
                xiaoHaoMenu.Enabled = false;

                //休风
                c1FlexGrid_xiufeng.AllowAddNew = false;
                c1FlexGrid_xiufeng.AllowDelete = false;
                c1FlexGrid_xiufeng.AllowEditing = false;
                xfMenu.Enabled = false;

                //工作记录
                ybTxt.ReadOnly = true;
                ybTxt.BackColor = Color.LightYellow;
                bbTxt.ReadOnly = true;
                bbTxt.BackColor = Color.LightYellow;
                zbTxt.ReadOnly = true;
                zbTxt.BackColor = Color.LightYellow;
                //记事本
                jsbTxt.ReadOnly = true;
                jsbTxt.BackColor = Color.LightYellow;
            }
        }

        private void FlexGrid_BeforeDeleteRow(object sender, C1.Win.C1FlexGrid.RowColEventArgs e)
        {
            if (MessageBox.Show("您确实要删除记录吗？", "询问", MessageBoxButtons.YesNo) == DialogResult.No)
                e.Cancel = true;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            quanxian();
            this.dateTimePicker2.Visible = false;
         //   this.label11.Visible = false;
            switch (操作参数平均.SelectedTab.Name)
            {

                case "SLQK":    //上料情况
                    this.rbliaoxianbbTableAdapter1.Fill(this.生产数据1.RBLIAOXIANBB, this.dateTimePicker1.Value.Date, Convert.ToDecimal(this.comboBox1.Text));
                    this.rbliaoxianybTableAdapter1.Fill(this.生产数据1.RBLIAOXIANYB, this.dateTimePicker1.Value.Date - TimeSpan.FromHours(1), this.dateTimePicker1.Value.Date + TimeSpan.FromHours(7), Convert.ToDecimal(this.comboBox1.Text));
                    this.rbliaoxianzbTableAdapter1.Fill(this.生产数据1.RBLIAOXIANZB, this.dateTimePicker1.Value.Date, Convert.ToDecimal(this.comboBox1.Text));
                    c1FlexGrid_yeban.Refresh();
                    c1FlexGrid_baiban.Refresh();
                    c1FlexGrid_zhongban.Refresh();
                    break;
                case "TZQK":    //铁渣情况
                    this.ddluciTableAdapter1.Fill(this.生产数据1.DDLUCI, this.dateTimePicker1.Value.Date, Convert.ToDecimal(this.comboBox1.Text)); //渣铁 情况
                     c1FlexGrid_tiezha.Refresh();
                     this.c1FlexGrid_tiezha.Subtotal(AggregateEnum.Clear);
                     c1FlexGrid_tiezha.Subtotal(AggregateEnum.Sum, 0, -1, 4, "合计值");
                     c1FlexGrid_tiezha.Subtotal(AggregateEnum.Sum, 0, -1, 7, "合计值");
                     c1FlexGrid_tiezha.Subtotal(AggregateEnum.Sum, 0, -1, 21, "合计值");
                     c1FlexGrid_tiezha.Subtotal(AggregateEnum.Average, 0, -1, 9, "平均值");


                     for (int i = 12; i < 19; i++)
                     {
                         c1FlexGrid_tiezha.Subtotal(AggregateEnum.Average, 0, -1, i, "平均值");

                     }
                     for (int j = 25; j < 34; j++)
                     {
                         c1FlexGrid_tiezha.Subtotal(AggregateEnum.Average, 0, -1, j, "平均值");

                     }
                     for (int c = 37; c < 48; c++)
                     {
                         c1FlexGrid_tiezha.Subtotal(AggregateEnum.Average, 0, -1, c, "平均值");

                     }
                     c1FlexGrid_tiezha.SubtotalPosition = SubtotalPositionEnum.BelowData;
                     CellStyle ss = c1FlexGrid_tiezha.Styles[CellStyleEnum.Subtotal0];
                     ss.BackColor = Color.WhiteSmoke;
                     ss.ForeColor = Color.Black;
                     ss.Font = new Font(c1FlexGrid_tiezha.Font, FontStyle.Bold);
                     ss.Format = "#######0.##";
                    break;
                case "BDQK":    //变动情况
                    this.rbbdTableAdapter1.Fill(this.生产数据1.RBBD, this.dateTimePicker1.Value.Date, Convert.ToDecimal(this.comboBox1.Text));
                    c1FlexGrid_biandong.Refresh();
                    break;
                case "CZCS":    //操作参数
                    this.rbcaozuoTableAdapter1.Fill(this.生产数据1.RBCAOZUO, this.dateTimePicker1.Value.Date, Convert.ToDecimal(this.comboBox1.Text)); //操作情况
                    c1FlexGrid_caozuo.Refresh();

                    this.c1FlexGrid_caozuo.Subtotal(AggregateEnum.Clear);
                    for (int c = 3; c < c1FlexGrid_caozuo.Cols.Count; c++)
                    {
                        c1FlexGrid_caozuo.Subtotal(AggregateEnum.Average, 0, -1, c, "平均值");
                      

                    }
                    c1FlexGrid_caozuo.SubtotalPosition = SubtotalPositionEnum.BelowData;
                    CellStyle s1 = c1FlexGrid_caozuo.Styles[CellStyleEnum.Subtotal0];
                    s1.BackColor = Color.WhiteSmoke;
                    s1.ForeColor = Color.Black;
                    s1.Font = new Font(c1FlexGrid_caozuo.Font, FontStyle.Bold);
                    s1.Format = "#######0.##";
                    break;
                case "ZLZD":    //装料制度
                    this.rbzlbdTableAdapter1.Fill(this.生产数据1.RBZLBD, this.dateTimePicker1.Value.Date, Convert.ToDecimal(this.comboBox1.Text)); //装料变动
                    c1FlexGrid_zhuangliao.Refresh();
                    break;
                case "FKQK":    //风口情况
                    this.rbfkTableAdapter1.Fill(this.生产数据1.RBFK, this.dateTimePicker1.Value.Date, Convert.ToDecimal(this.comboBox1.Text)); //风口情况
                    c1FlexGrid_fengkou.Refresh();
                    break;
                case "LQQK":
                    this.rblqTableAdapter1.Fill(this.生产数据1.RBLQ, this.dateTimePicker1.Value.Date, Convert.ToDecimal(this.comboBox1.Text)); //冷却情况
                 
                    
                    C1FlexGrid flex = c1FlexGrid_lengque as C1FlexGrid;
                    if (flex.Rows.Count > 1)
                    {
                        for (int c = 1; c < flex.Cols.Count; c++)
                        {
                            if (c < flex.Cols.Fixed || flex.Cols[c].DataType == typeof(System.String) || flex.Cols[c].DataType == typeof(System.DateTime) || flex.Cols[c].DataType == typeof(System.Boolean) || flex.Cols[c].DataType == typeof(System.Byte) || flex.Cols[c].DataType == typeof(System.Char) || flex.Cols[c].DataType == typeof(System.SByte) || flex.Cols[c].DataType == typeof(System.Object))
                                continue;
                            flex.Subtotal(AggregateEnum.Average, 0, -1, c, "平均值");
                            if (flex.Cols[c] != null && (flex.Cols[c].DataType == typeof(System.Int32) || flex.Cols[c].DataType == typeof(System.Decimal)))
                            {
                                decimal de = string.IsNullOrEmpty(Convert.ToString(flex.GetData(flex.Rows.Count - 1, c))) ? 0 : Convert.ToDecimal(flex.GetData(flex.Rows.Count - 1, c));
                                flex.SetData(flex.Rows.Count - 1, c, de.ToString("0.00"));
                            }
                        }
                  
                        CellStyle s3 = flex.Styles[CellStyleEnum.Subtotal0];
                        s3.BackColor = Color.WhiteSmoke;
                        s3.ForeColor = Color.Black;
                        s3.Font = new Font(flex.Font, FontStyle.Bold);
                        s3.Format = "#######0.##";
                    }
                    c1FlexGrid_lengque.Refresh();
                   // c1FlexGrid_lengque.SubtotalPosition = SubtotalPositionEnum.BelowData;

                    break;
                case "XHQK":    //消耗情况
                    this.rbxiaohaoTableAdapter1.Fill(this.生产数据1.RBXIAOHAO, this.dateTimePicker1.Value.Date, Convert.ToDecimal(this.comboBox1.Text)); //消耗情况
                    c1FlexGrid_xiaohao.Refresh();
                    break;
                case "XFFF":    //休风放风
                    rbxfTable.LoadByRQGAOLU(this.dateTimePicker1.Value.Date, Convert.ToInt32(this.comboBox1.Text));
                   
                    //this.rbxfTableAdapter1.Fill(this.生产数据1.RBXF, this.dateTimePicker1.Value.Date, Convert.ToDecimal(this.comboBox1.Text)); //休风情况
                   // c1FlexGrid_xiufeng.Refresh();
                    break;
            
                case "GZRZ":
                    OracleConnection cn = new OracleConnection();
                    cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
                    cn.Open();
                    OracleCommand cmd = new OracleCommand();
                    cmd.Connection = cn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "SELECT JL FROM RBGZ WHERE TRUNC(SJ)=:RQ and gaolu=:GAOLU and bc='夜班'";
                    cmd.Parameters.Add(":RQ", OracleType.DateTime).Value = this.dateTimePicker1.Value.Date;
                    cmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = Convert.ToInt32(this.comboBox1.Text);
                    OracleDataReader dr = cmd.ExecuteReader();
                    if (dr.Read())
                        ybTxt.Text = dr.IsDBNull(0) ? "" : dr.GetString(0);
                    else
                        ybTxt.Text = "";
                    dr.Close();
                    cmd.CommandText = "SELECT JL FROM RBGZ WHERE TRUNC(SJ)=:RQ and gaolu=:GAOLU and bc='白班'";
                    cmd.Parameters.Add(":RQ", OracleType.DateTime).Value = this.dateTimePicker1.Value.Date;
                    cmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = Convert.ToInt32(this.comboBox1.Text);
                    dr = cmd.ExecuteReader();
                    if (dr.Read())
                        bbTxt.Text = dr.IsDBNull(0) ? "" : dr.GetString(0);
                    else
                        bbTxt.Text = "";
                    dr.Close();
                    cmd.CommandText = "SELECT JL FROM RBGZ WHERE TRUNC(SJ)=:RQ and gaolu=:GAOLU and bc='中班'";
                    cmd.Parameters.Add(":RQ", OracleType.DateTime).Value = this.dateTimePicker1.Value.Date;
                    cmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = Convert.ToInt32(this.comboBox1.Text);
                    dr = cmd.ExecuteReader();
                    if (dr.Read())
                        zbTxt.Text = dr.IsDBNull(0) ? "" : dr.GetString(0);
                    else
                        zbTxt.Text = "";
                    dr.Close();
                    cn.Close();
                    break;
                case "JSB":

                    OracleConnection cn1 = new OracleConnection();
                    cn1.ConnectionString = Properties.Settings.Default.ltznConnectionString;
                    cn1.Open();
                    OracleCommand cmd1 = new OracleCommand();
                    cmd1.Connection = cn1;
                    cmd1.CommandType = CommandType.Text;
                    cmd1.CommandText = "SELECT JS FROM RBJS WHERE TRUNC(SJ)=:RQ and gaolu=:GAOLU";
                    cmd1.Parameters.Add(":RQ", OracleType.DateTime).Value = this.dateTimePicker1.Value.Date;
                    cmd1.Parameters.Add(":GAOLU", OracleType.Int32).Value = Convert.ToInt32(this.comboBox1.Text);
                    dr = cmd1.ExecuteReader();
                    if (dr.Read())
                        jsbTxt.Text = dr.IsDBNull(0) ? "" : dr.GetString(0);
                    else
                        jsbTxt.Text = "";
                    dr.Close();
                    cn1.Close();
                    break;
                case "CZCSPJZ":    //操作参数
                    this.dateTimePicker2.Visible = true;
                  //  this.label11.Visible = true;
                    this.rbcaozuO1TableAdapter1.Fill(this.生产数据1.RBCAOZUO1, this.dateTimePicker2.Value.Date, this.dateTimePicker1.Value.Date, Convert.ToDecimal(this.comboBox1.Text)); //操作情况
      
                    c1FlexGrid_caozuopjz.Refresh();

                    this.c1FlexGrid_caozuopjz.Subtotal(AggregateEnum.Clear);
                    for (int c = 3; c < c1FlexGrid_caozuopjz.Cols.Count; c++)
                    {
                        c1FlexGrid_caozuopjz.Subtotal(AggregateEnum.Average, 0, -1, c, "平均值");
                    }
                    c1FlexGrid_caozuopjz.SubtotalPosition = SubtotalPositionEnum.BelowData;
                    CellStyle s2 = c1FlexGrid_caozuopjz.Styles[CellStyleEnum.Subtotal0];
                    s2.BackColor = Color.WhiteSmoke;
                    s2.ForeColor = Color.Black;
                    s2.Font = new Font(c1FlexGrid_caozuopjz.Font, FontStyle.Bold);
                    s2.Format = "#######0.##";
                    break;
                case "TSCF":    //成分参数
                    //this.dateTimePicker2.Visible = true;
                    //  this.label11.Visible = true;
                    this.cqA_IRON_ELEMENT_TEXTTableAdapter1.Fill(this.生产数据1.CQA_IRON_ELEMENT_TEXT, this.dateTimePicker1.Value.Date); //铁水成分查询
                    dataGridView_chengfen.Refresh();
                    break;

            }
           // button1.Enabled = false;
            this.Cursor = Cursors.Default;
        }

        private void c1FlexGrid_tiezha_CellChanged(object sender, RowColEventArgs e)
        {
            DateTime OldDukouSj = this.dateTimePicker1.Value;         //上个堵口时间
            DateTime DukouSj = this.dateTimePicker1.Value; ;         //堵口时间
            Decimal MPiShu;             //到上个装料变动批数
            Decimal NPiShu;             //到下个装料变动批数
            Decimal TiePianCha = 0;        //亏铁吨数
            DateTime ZlbdSjNew = this.dateTimePicker1.Value;         //最新的装料变动时间
            Decimal ZlbdTieLNew = 0;          //最新的装料的铁量
            DateTime ZlbdSjOld = this.dateTimePicker1.Value;         //旧的装料变动时间
            Decimal ZlbdTieLOld = 0;          //旧的装料的铁量
            Decimal ZongLLTL = 0;           //总的理论铁量;

          
            if (!(c1FlexGrid_tiezha[e.Row, 3] is DBNull))               //当堵口时间过来了
            {
                //总料批数
                decimal PiShu = c1FlexGrid_tiezha[e.Row, "LIAOPI"] != DBNull.Value ? Convert.ToDecimal(c1FlexGrid_tiezha[e.Row, "LIAOPI"]) : 0; ;
                if (e.Row > 1 && c1FlexGrid_tiezha[e.Row, "DKSJ"] != DBNull.Value)
                {
                    for (int i = e.Row - 1; i > 0; i--)
                    {
                        if (c1FlexGrid_tiezha[i, "FELIANG"] != DBNull.Value && c1FlexGrid_tiezha[i, "LFELIANG"] != DBNull.Value && c1FlexGrid_tiezha[i, 3] != DBNull.Value)
                        {
                            TiePianCha = (decimal)c1FlexGrid_tiezha[i, "FELIANG"] - (decimal)c1FlexGrid_tiezha[i, "LFELIANG"];
                            OldDukouSj = (DateTime)c1FlexGrid_tiezha[i, 3];
                            break;
                        }
                    }
                }

                //打开装料制度表
                OracleConnection cn = new OracleConnection();
                cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
                cn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = cn;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = "select blsj,lltl from rbzlbd where gaolu =:gaolu and trunc(sj) =:sj order by sj";
                cmd.Parameters.Add(":sj", OracleType.DateTime).Value = this.dateTimePicker1.Value.Date;
                cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = Convert.ToInt32(this.comboBox1.Text);
                OracleDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    if (ZlbdSjNew != null && !dr.IsDBNull(0))
                    {
                        ZlbdSjOld = ZlbdSjNew;
                    }
                    if (!dr.IsDBNull(1))
                    {
                        ZlbdTieLOld = ZlbdTieLNew;

                    }
                    if (!dr.IsDBNull(0))
                    {
                        ZlbdSjNew = dr.GetDateTime(0);
                        ZlbdSjNew = new DateTime(this.dateTimePicker1.Value.Year, this.dateTimePicker1.Value.Month, this.dateTimePicker1.Value.Day, ZlbdSjNew.Hour, ZlbdSjNew.Minute, ZlbdSjNew.Second);
                    }
                    if (!dr.IsDBNull(1))
                    {
                        ZlbdTieLNew = dr.GetDecimal(1);

                    }

                }
                if (c1FlexGrid_tiezha[e.Row, 3] != DBNull.Value)
                {
                    DukouSj = (DateTime)c1FlexGrid_tiezha[e.Row, 3];      //当前堵口时间
                }

                if (PiShu > 0 && ZlbdTieLNew != ZlbdTieLOld && ZlbdTieLOld != 0)           //判断是否有料批数，是否有装料变动
                {
                    if (OldDukouSj < ZlbdSjNew && DukouSj > ZlbdSjNew)
                    {
                        OracleCommand cmd1 = new OracleCommand();
                        cmd1.CommandText = "select count(*) from lt_liao where gaolu = :gaolu and T >=:sj1 and T <=:sj2";
                        cmd1.Connection = cn;
                        cmd1.Parameters.Add(":gaolu", OracleType.Int32).Value = Convert.ToInt32(this.comboBox1.Text);
                        cmd1.Parameters.Add(":sj1", OracleType.DateTime).Value = OldDukouSj;
                        cmd1.Parameters.Add(":sj2", OracleType.DateTime).Value = ZlbdSjNew;
                        MPiShu = (Decimal)cmd1.ExecuteScalar();
                        NPiShu = PiShu - MPiShu;
                        ZongLLTL = MPiShu * ZlbdTieLOld / 1000 + NPiShu * ZlbdTieLNew / 1000;

                    }
                    else
                    {

                        ZongLLTL = PiShu * ZlbdTieLNew / 1000;

                    }
                }
                else
                {
                    ZongLLTL = PiShu * ZlbdTieLNew / 1000;
                }

                if (TiePianCha < 0)
                {
                    ZongLLTL = ZongLLTL - TiePianCha;
                }

                cn.Close();

                c1FlexGrid_tiezha[e.Row, "LFELIANG"] = ZongLLTL;
            }

        }

        private void c1FlexGrid_caozuo_AfterAddRow(object sender, RowColEventArgs e)
        {

            if (c1FlexGrid_caozuo[e.Row, "HOUR"] == DBNull.Value)
            {
                c1FlexGrid_caozuo[e.Row, "SJ"] = this.dateTimePicker1.Value.Date;
                c1FlexGrid_caozuo[e.Row, "GAOLU"] = Convert.ToInt32(this.comboBox1.Text);

            }
            else
            {
                int t = Convert.ToInt32(c1FlexGrid_caozuo[e.Row, "HOUR"]);
                int year = this.dateTimePicker1.Value.Year;
                int month = this.dateTimePicker1.Value.Month;
                int day = this.dateTimePicker1.Value.Day;
                c1FlexGrid_caozuo[e.Row, "SJ"] = new DateTime(year, month, day, t, 0, 0);
                c1FlexGrid_caozuo[e.Row, "GAOLU"] = Convert.ToInt32(this.comboBox1.Text);
            }
        }
     
        private void 生产日报_Load(object sender, EventArgs e)
        {
            this.c1FlexGrid_xiufeng.DataSource = rbxfTable;
            rbxfTable.ListChanged += new ListChangedEventHandler(chutieSource_ListChanged);

        }
        void chutieSource_ListChanged(object sender, ListChangedEventArgs e)
        {

            btnSave.Enabled = true;
        }
        private void c1FlexGrid_zhuangliao_AfterAddRow(object sender, RowColEventArgs e)
        {
            c1FlexGrid_zhuangliao[e.Row, "SJ"] = this.dateTimePicker1.Value.Date;
            c1FlexGrid_zhuangliao[e.Row, "GAOLU"] = Convert.ToInt32(this.comboBox1.Text);
        }

        private void c1FlexGrid_zhuangliao_AfterEdit(object sender, RowColEventArgs e)
        {
            C1FlexGrid flex = sender as C1FlexGrid;

            if (flex == null || e.Col < 0)
            {
                
                return;
            }
            string colName = flex.Cols[e.Col].Name;

            //计算理论焦比
            if (colName == "LLTL" || colName == "JP" || colName == "JD")
            {
                double lltl = flex[e.Row, "LLTL"] != DBNull.Value ? Convert.ToDouble(flex[e.Row, "LLTL"]) : 0;
                double jp = flex[e.Row, "JP"] != DBNull.Value ? Convert.ToDouble(flex[e.Row, "JP"]) : 0;
                double jd = flex[e.Row, "JD"] != DBNull.Value ? Convert.ToDouble(flex[e.Row, "JD"]) : 0;
                if (lltl > 0)
                {
                    double lljb = (jp + jd) * 1000 / lltl;
                    flex[e.Row, "LLJB"] = lljb.ToString("##0.00");
                }
                else
                {
                    flex[e.Row, "LLJB"] = DBNull.Value;
                }
                flex.Refresh();
                lltl = 0;
                jp = 0;
                jd = 0;

            }

            //计算理论班产
            if (colName == "LLTL" || colName == "YQ")
            {
                double lltl = flex[e.Row, "LLTL"] != DBNull.Value ? Convert.ToDouble(flex[e.Row, "LLTL"]) : 0;
                double pici = flex[e.Row, "PICI"] != DBNull.Value ? Convert.ToDouble(flex[e.Row, "PICI"]) : 0;
                double yq = flex[e.Row, "YQ"] != DBNull.Value ? Convert.ToDouble(flex[e.Row, "YQ"]) : 0;

                double llbc = lltl * yq / 1000;
                flex[e.Row, "LLBC"] = llbc.ToString("##0.00");
                flex.Refresh();
                lltl = 0;
                pici = 0;
                yq = 0;
            }

            //计算综合负荷
            if (colName == "KP" || colName == "JD" || colName == "JP" || colName == "PML" || colName == "YQ")
            {

                double kuang = flex[e.Row, "KP"] != DBNull.Value ? Convert.ToDouble(flex[e.Row, "KP"]) : 0;
                double jiao1 = flex[e.Row, "JD"] != DBNull.Value ? Convert.ToDouble(flex[e.Row, "JD"]) : 0;
                double jiao2 = flex[e.Row, "JP"] != DBNull.Value ? Convert.ToDouble(flex[e.Row, "JP"]) : 0;
                double mei = flex[e.Row, "PML"] != DBNull.Value ? Convert.ToDouble(flex[e.Row, "PML"]) : 0;
                double yq = flex[e.Row, "YQ"] != DBNull.Value ? Convert.ToDouble(flex[e.Row, "YQ"]) : 0;

                double jiao = jiao1 + jiao2 + mei * 900 * 8 / yq;

                if (jiao > 0 && yq > 0)
                {
                    double zhfh = kuang / jiao;
                    flex[e.Row, "ZHFH"] = (zhfh).ToString("##0.00");
                }
                else
                {
                    flex[e.Row, "ZHFH"] = DBNull.Value;
                }
                flex.Refresh();
                kuang = 0;
                jiao1=0;
                jiao2 = 0;
                mei = 0;
                yq = 0;

            }
            //计算焦炭负荷
            if (colName == "KP" || colName == "JP")
            {

                double kuang = flex[e.Row, "KP"] != DBNull.Value ? Convert.ToDouble(flex[e.Row, "KP"]) : 0;
                double jiao2 = flex[e.Row, "JP"] != DBNull.Value ? Convert.ToDouble(flex[e.Row, "JP"]) : 0;
                double jiao1 = flex[e.Row, "JD"] != DBNull.Value ? Convert.ToDouble(flex[e.Row, "JD"]) : 0;
                if (jiao2 + jiao1 > 0)
                {
                    double jtfh = kuang / (jiao2 + jiao1);
                    flex[e.Row, "JTFH"] = (jtfh).ToString("##0.00");
                }
                else
                {
                    flex[e.Row, "JTFH"] = DBNull.Value;
                }
                flex.Refresh();
                kuang = 0;
                jiao1 = 0;
                jiao2 = 0;
            }
        }


        private void button2_Click(object sender, EventArgs e)
        {
            OracleConnection cn = new OracleConnection();
            cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
            cn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = cn;
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "INSERT INTO RBZLBD(GAOLU, SJ, BC, PICI, JS, SQ, QT, JT, JD, KP, JP, JTFH, PML, ZHFH, ZL, R2, LLJD, ZLSX, LX, BLSJ, BLYY, CLCH, LLTL, BZ, YQ, LLBC, BLJD) Select * from (Select GAOLU, SYSDATE, BC, PICI, JS, SQ, QT, JT, JD, KP, JP, JTFH, PML, ZHFH, ZL, R2, LLJD, ZLSX, LX, BLSJ, BLYY, CLCH, LLTL, BZ, YQ, LLBC, BLJD FROM RBZLBD where TRUNC(SJ)>=:SJ AND GAOLU=:GAOLU  ORDER BY SJ DESC) WHERE ROWNUM=1";
            cmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = Convert.ToInt32(this.comboBox1.Text);
            cmd.Parameters.Add(":SJ", OracleType.DateTime).Value = this.dateTimePicker1.Value.Date.AddDays(-1);
            cmd.ExecuteScalar();
            cn.Close();
            this.rbzlbdTableAdapter1.Fill(this.生产数据1.RBZLBD, this.dateTimePicker1.Value.Date, Convert.ToDecimal(this.comboBox1.Text)); //装料变动
            c1FlexGrid_zhuangliao.Refresh();
        }

       

        private void c1FlexGrid_biandong_AfterEdit(object sender, RowColEventArgs e)
        {
            if (!(c1FlexGrid_biandong[e.Row, 3] is DBNull))
            {
                string tiaojian;
                tiaojian = (string)c1FlexGrid_biandong[e.Row, 3];
                if (e.Col == 3)
                {
                    switch (tiaojian)
                    {
                        case "加风":
                            c1FlexGrid_biandong[e.Row, 5] = "kpa";
                            break;
                        case "减风":
                            c1FlexGrid_biandong[e.Row, 5] = "kpa";
                            break;
                        case "加煤":
                            c1FlexGrid_biandong[e.Row, 5] = "t/h";
                            break;
                        case "减煤":
                            c1FlexGrid_biandong[e.Row, 5] = "t/h";
                            break;
                        case "加氧":
                            c1FlexGrid_biandong[e.Row, 5] = "m3/h";
                            break;
                        case "减氧":
                            c1FlexGrid_biandong[e.Row, 5] = "m3/h";
                            break;
                    }
                }

            }
        }

        private void c1FlexGrid_biandong_AfterAddRow(object sender, RowColEventArgs e)
        {

            c1FlexGrid_biandong[e.Row, "RQ"] = this.dateTimePicker1.Value.Date;
            c1FlexGrid_biandong[e.Row, "GAOLU"] = Convert.ToInt32(this.comboBox1.Text);

        }

        /// <summary>
        /// 消耗情况计算
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void c1FlexGrid_xiaohao_AfterEdit(object sender, RowColEventArgs e)
        {
            C1FlexGrid flex = sender as C1FlexGrid;

            if (flex == null || e.Col < 0)
            {
                return;
            }
            string colName = flex.Cols[e.Col].Name;

            if (colName == "YEBAN" || colName == "BAIBAN" || colName == "ZHONGBAN")
            {
                double yeban = flex[e.Row, "YEBAN"] != DBNull.Value ? Convert.ToDouble(flex[e.Row, "YEBAN"]) : 0;
                double baiban = flex[e.Row, "BAIBAN"] != DBNull.Value ? Convert.ToDouble(flex[e.Row, "BAIBAN"]) : 0;
                double zhongban = flex[e.Row, "ZHONGBAN"] != DBNull.Value ? Convert.ToDouble(flex[e.Row, "ZHONGBAN"]) : 0;
                c1FlexGrid_xiaohao[e.Row, "LEIJI"] = yeban + baiban + zhongban;
            }
        }

        ///// <summary>
        ///// 休风计算
        ///// </summary>
        ///// <param name="sender"></param>
        ///// <param name="e"></param>
        //private void c1FlexGrid_xiufeng_AfterEdit(object sender, RowColEventArgs e)
        //{

        //    C1FlexGrid flex = sender as C1FlexGrid;

        //    if (flex == null || e.Col < 0)
        //    {
        //        return;
        //    }
        //    string colName = flex.Cols[e.Col].Name;

        //    if (colName == "KS" || colName == "ZZ")
        //    {
        //        DateTime rq = this.dateTimePicker1.Value.Date;

        //        DateTime? ks = flex[e.Row, "KS"] != DBNull.Value ? (DateTime?)new DateTime(rq.Year, rq.Month, rq.Day, ((DateTime)flex[e.Row, "KS"]).Hour, ((DateTime)flex[e.Row, "KS"]).Minute, 0) : null;
        //        DateTime? zz = flex[e.Row, "ZZ"] != DBNull.Value ? (DateTime?)new DateTime(rq.Year, rq.Month, rq.Day, ((DateTime)flex[e.Row, "ZZ"]).Hour, ((DateTime)flex[e.Row, "ZZ"]).Minute, 0) : null;
        //        TimeSpan? jg = zz - ks;
        //        if (jg != null)
        //            flex[e.Row, "JG"] = jg.Value.TotalMinutes;
        //        else
        //            flex[e.Row, "JG"] = DBNull.Value;

        //    }
        //}

      

        private void 删除ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            deleteFlexGridRow(c1FlexGrid_xiaohao);
        }

        private void deleteFlexGridRow(C1FlexGrid flex)
        {
            if (MessageBox.Show("您确实要删除记录吗？", "询问", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                if (flex != null)
                {
                    CurrencyManager cm = (CurrencyManager)BindingContext[flex.DataSource, flex.DataMember];
                    cm.EndCurrentEdit();
                    if (flex.RowSel > 0)
                    {
                        flex.Rows.Remove(flex.Row);
                        btnSave.Enabled = true;
                    }
                }
            }
        }

        private void xfDeleteMenuItem_Click(object sender, EventArgs e)
        {
            deleteFlexGridRow(c1FlexGrid_xiufeng);
        }

        private void lqDeleteMenuItem_Click(object sender, EventArgs e)
        {
            deleteFlexGridRow(c1FlexGrid_lengque);
        }

        private void c1FlexGrid_ChangeEdit(object sender, EventArgs e)
        {
              btnSave.Enabled = true;
        }

         private void ybTxt_TextChanged(object sender, EventArgs e)
        {
            btnSave.Enabled = true;
        }

        private void fkMenuItem_Click(object sender, EventArgs e)
        {
            deleteFlexGridRow(c1FlexGrid_fengkou);
        }

           /// <summary>
        /// 操作参数的计算情况
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void c1FlexGrid_caozuo_AfterEdit(object sender, RowColEventArgs e)
        {
            double LFLL = 0;                 //冷风流量
            double FYLL = 0;                 //富氧流量
            double RFWD = 0;                 //热风温度
            double MEI = 0;                  //喷煤量
            DateTime shijian;              // 当前时间
            double kuang = 0;              //一小时上的矿
            double jiao = 0;              //一小时上的焦

            if (e.Col == 9)
            {
                shijian = this.dateTimePicker1.Value.Date;
                if (!(c1FlexGrid_caozuo[e.Row, "HOUR"] is DBNull))
                {
                    int t = Convert.ToInt32(c1FlexGrid_caozuo[e.Row, "HOUR"]);
                    int year = this.dateTimePicker1.Value.Year;
                    int month = this.dateTimePicker1.Value.Month;
                    int day = this.dateTimePicker1.Value.Day;
                    shijian = new DateTime(year, month, day, t, 0, 0);

                }
                if (!(c1FlexGrid_caozuo[e.Row, "LFLL"] is DBNull))
                {

                    LFLL = Convert.ToDouble(c1FlexGrid_caozuo[e.Row, "LFLL"]);
                }
                if (!(c1FlexGrid_caozuo[e.Row, "FYLL"] is DBNull))
                {
                    FYLL = Convert.ToDouble(c1FlexGrid_caozuo[e.Row, "FYLL"]);
                }
                if (!(c1FlexGrid_caozuo[e.Row, "RFWD"] is DBNull))
                {
                    RFWD = Convert.ToDouble(c1FlexGrid_caozuo[e.Row, "RFWD"]);
                }
                if (!(c1FlexGrid_caozuo[e.Row, "PML"] is DBNull))
                {
                    MEI = Convert.ToDouble(c1FlexGrid_caozuo[e.Row, "PML"]);
                }
                if (LFLL != 0)
                {
                    c1FlexGrid_caozuo[e.Row, "RSWD"] = (1690 + RFWD * 0.839 + FYLL / (6 * LFLL / 100) - 4.972 * MEI * 1000 / (6 * LFLL / 100)).ToString("0000");   //燃烧温度
                }
                else
                {
                    c1FlexGrid_caozuo[e.Row, "RSWD"] = 0;
                }
                if (Convert.ToInt32(c1FlexGrid_caozuo[e.Row, "RSWD"]) < 0)
                {
                    c1FlexGrid_caozuo[e.Row, "RSWD"] = 0;
                }

                if (LFLL != 0 && MEI != 0 && RFWD != 0 && FYLL != 0 && LFLL != 0)
                {
                    OracleConnection cn = new OracleConnection();
                    cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
                    cn.Open();
                    OracleCommand cmd = new OracleCommand();
                    cmd.Connection = cn;
                    cmd.CommandType = CommandType.Text;
                    cmd.CommandText = "select sum(sjk),sum(slq),sum(jt),sum(jd) from lt_liao where T< =:sj1 and T >=:sj2  and gaolu=:gaolu";
                    cmd.Parameters.Add(":sj1", OracleType.DateTime).Value = shijian;
                    cmd.Parameters.Add(":sj2", OracleType.DateTime).Value = shijian.AddHours(-1);
                    cmd.Parameters.Add(":gaolu", OracleType.Int32).Value = Convert.ToInt32(this.comboBox1.Text);
                    OracleDataReader dr = cmd.ExecuteReader();
                    while (dr.Read())
                    {
                        if (!dr.IsDBNull(0) && !dr.IsDBNull(1))
                        {
                            kuang = kuang + dr.GetDouble(0) + dr.GetDouble(1);
                        }
                        if (!dr.IsDBNull(2) && !dr.IsDBNull(3))
                        {
                            jiao = jiao + dr.GetDouble(2) + dr.GetDouble(3);
                        }
                    }
                    cn.Close();
                    if (jiao + MEI > 0)
                    {
                        c1FlexGrid_caozuo[e.Row, "ZHFH"] = (kuang / (jiao + MEI * 900)).ToString("##0.00");
                    }
                    else
                    {
                        c1FlexGrid_caozuo[e.Row, "ZHFH"] = 0;
                    }
                }
            }
        }

        private void zlzdDeleteMenuItem_Click(object sender, EventArgs e)
        {
            deleteFlexGridRow(c1FlexGrid_zhuangliao);
        }

        private void zlbdDeleteMenuItem_Click(object sender, EventArgs e)
        {
            deleteFlexGridRow(c1FlexGrid_biandong);
        }

        private void c1FlexGrid_lengque_AfterEdit(object sender, RowColEventArgs e)
        {
            C1FlexGrid flex = sender as C1FlexGrid;
            double lqsum = 0;
            int lqyxs=0;
            if (flex == null || e.Col < 0)
            {
                return;
            }
            string colName = flex.Cols[e.Col].Name;
            for(int i=5;i< 54; i++)
            {
               
               double lqzh = flex[e.Row, i] != DBNull.Value ? Convert.ToDouble(flex[e.Row, i]) : 0;
               if (lqzh > 0)
               {
                   lqyxs = lqyxs + 1;
                   lqsum = lqsum + lqzh;
               }
            }
            if (lqyxs > 0)
            {
                flex[e.Row, 4] = (lqsum / lqyxs).ToString("0.00");
            }
        }

        private void c1FlexGrid_lengque_AfterAddRow(object sender, RowColEventArgs e)
        {
            c1FlexGrid_lengque[e.Row, "SJ"] = System.DateTime.Now;
            c1FlexGrid_lengque[e.Row, "GAOLU"] = Convert.ToInt32(this.comboBox1.Text);
           
        }

        private void c1FlexGrid_caozuo_Click(object sender, EventArgs e)
        {

        }

        private void TSCF_Click(object sender, EventArgs e)
        {

        }

       
    }
}