using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;
using Microsoft.Reporting.WinForms;
using LTZN.Repository;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Runtime.Serialization;
using Microsoft.Office.Interop.Excel;

namespace LTZN.技术日报
{
    public partial class RiReportForm : Form
    {
        public RiReportForm()
        {
            InitializeComponent();
            LtznUserManager.instance.RegisterHandler(this, UserChange);
        }

        //用户变动时
        void UserChange(LtznUser ltznUser)
        {
            if (ltznUser != null && ltznUser.IsInRole("统计"))
            {
                btnQuery.Enabled = true;
                btnSave.Enabled = true;
            }
            else
            {
                btnQuery.Enabled = false;
                btnSave.Enabled = false;

            }

            if (ltznUser != null && ltznUser.IsInRole("技术日报模型"))
            {
                btnModel.Enabled = true;
             
            }
            else
            {
                btnModel.Enabled = false;
            }
        }

        private 技术日报 _jiShuRiBao = null;
        public 技术日报 JiShuRiBao
        {
            get {
                if (_jiShuRiBao == null)
                {
                    _jiShuRiBao = new 技术日报();
                }
                return _jiShuRiBao;
            }
        }
      
        private void RiReportForm_Load(object sender, EventArgs e)
        {
      
            webBrowser1.DocumentText = "<HTML><HEAD><style type='text/css'> table{ border:1px; border-cllapse:collapse;margin:0;padding:0;}  td{border:thin solid red;border-cllapse:collapse;padding:0;}</style></HEAD><body bgcolor=#EEEEFF topmargin=25 leftmargin=5>&nbsp;</BODY></HTML>";
            riqiControl.Value = DateTime.Today;
        }
        //生成日报
        private void btnQuery_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            this.reportViewer1.Visible = false;
            this.JiShuRiBao.myModel.Init();
            this.JiShuRiBao.getDataBy(riqiControl.Value.Date);
            this.JiShuRiBao.ModelCalc();
            this.JiShuRiBao.display(webBrowser1.Document);
            this.Cursor = Cursors.Default;
        }
        /// <summary>
        /// 显示日报计算模型
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnModel_Click(object sender, EventArgs e)
        {
            if (this.JiShuRiBao.myModel != null)
            {
                CalModelForm form = new CalModelForm(this.JiShuRiBao.myModel);
                form.ShowDialog();
                this.JiShuRiBao.display(webBrowser1.Document);
            }
        }
        /// <summary>
        ///  打印
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPrint_Click(object sender, EventArgs e)
        {
            if (reportViewer1.Visible)
                reportViewer1.PrintDialog();
            else
                webBrowser1.ShowPrintPreviewDialog();
        }

        /// <summary>
        /// 显示报表
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnShowReport_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            this.reportViewer1.Visible = true;
            RiReportRepository rp = new RiReportRepository();
            技术日报内容 n = rp.GetRiReport(riqiControl.Value.Date);

            //技术日报数据集 ds = new 技术日报数据集();
            //this.技术日报TableAdapter1.FillByRq(ds.技术日报, riqiControl.Value.Date);
            ////报表显示窗体 f = new 报表显示窗体();
            ////f.WindowState = FormWindowState.Maximized;
            ////f.setDataSource(ds.技术日报,riqiControl.Value.Date);
            ////f.Show();
            //setDataSource(ds.技术日报, riqiControl.Value.Date);
            setDataSource(n);
            this.Cursor = Cursors.Default;
        }
        /// <summary>
        /// 保存
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSave_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            DateTime riqi=riqiControl.Value.Date;
            List<string> items=new List<string>();
            OracleConnection conn=new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            OracleCommand cmd = new OracleCommand("Select P01单位,P02项目 from 技术日报  Where (P日期 = :RQ)", conn);
            cmd.Parameters.Add(":RQ",OracleType.DateTime).Value=riqi;
            OracleDataReader dr=cmd.ExecuteReader();
            while(dr.Read())
            {
                if(!dr.IsDBNull(0) && !dr.IsDBNull(1))
                {
                    items.Add(dr.GetString(0) +"-"+dr.GetString(1));
                }
            }
            dr.Close();
            conn.Close();

            List<string> xiangmuList = new List<string>() { "本日", "累计" };
            for (int gaolu = 1; gaolu < 7; gaolu++)
            {
                if (gaolu == 2 || gaolu == 4) continue;
                foreach (var xiangmu in xiangmuList)
                {

                    if (items.Contains(gaolu.ToString() + "#-" + xiangmu))
                        this.技术日报TableAdapter1.Update(
                            (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.合格铁", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.炼钢铁", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.铸造铁", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.号外铁", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.合格率", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.利用系数", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.实物系数", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.原料矿总耗", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.原料矿单耗", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.机烧", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.竖球", xiangmu, gaolu))
                                                            , 0
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.其它熟料", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.本溪矿", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.其它生料", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.废铁总耗", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.废铁单耗", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.回收率", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.熟料比", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.平均风温", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.炉顶温度", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.热风压力", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.炉顶压力", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.富氧率", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.入炉焦炭总耗", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.入炉焦炭单耗", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.煤粉总耗", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.煤粉单耗", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.焦丁总耗", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.焦丁单耗", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.综合焦炭总耗", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.综合焦炭单耗", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.综合折算焦比", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.冶炼强度", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.焦炭负荷", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.干毛焦", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.炼钢铁Si", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.炼钢铁Mn", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.炼钢铁P", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.炼钢铁S", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.铸造铁Si", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.铸造铁Mn", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.铸造铁P", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.铸造铁S", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.炉渣碱度", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.休风情况", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.慢风", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.坐料次数", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.悬料次数", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.崩料次数", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.风口损坏数大", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.风口损坏数中", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.风口损坏数小", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.渣口损坏数大", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.渣口损坏数中", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.渣口损坏数小", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.本厂铸造SI大", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.本厂铸造SI小", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.送炼钢厂SI大", xiangmu, gaolu))
                                                            , 0
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.送炼钢厂SI小", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.折算产量", xiangmu, gaolu))
                                                            , 1
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.工艺称焦比", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.S小于002", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.P小于009", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.TI小于055", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.PB块", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.纽曼块", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.钛球", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.锰矿", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.硅石", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.白云石", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.蛇纹石", xiangmu, gaolu))

                                                             //, (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.球团矿", xiangmu, gaolu))
                                                             // , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.国内球团矿", xiangmu, gaolu))
                                                             //  , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.进口球团矿", xiangmu, gaolu))
                                                             //   , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.其它块矿", xiangmu, gaolu))
                                                             //    , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.高钛球团矿", xiangmu, gaolu))
                                                             //     , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.高品位钛球", xiangmu, gaolu))
                                                             //      , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.其它熔剂", xiangmu, gaolu))


                                                                , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.萤石", xiangmu, gaolu))
                                                                                                                      ,riqi, string.Format("{0}#", gaolu), xiangmu);
                    else
                        this.技术日报TableAdapter1.Insert(riqi,
                                                             string.Format("{0}#", gaolu),
                                                             xiangmu
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.合格铁", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.炼钢铁", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.铸造铁", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.号外铁", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.合格率", xiangmu, gaolu))
                                                            , (decimal?)(decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.利用系数", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.实物系数", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.原料矿总耗", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.原料矿单耗", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.机烧", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.竖球", xiangmu, gaolu))
                                                            , 0
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.其它熟料", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.本溪矿", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.其它生料", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.废铁总耗", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.废铁单耗", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.回收率", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.熟料比", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.平均风温", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.炉顶温度", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.热风压力", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.炉顶压力", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.富氧率", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.入炉焦炭总耗", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.入炉焦炭单耗", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.煤粉总耗", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.煤粉单耗", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.焦丁总耗", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.焦丁单耗", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.综合焦炭总耗", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.综合焦炭单耗", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.综合折算焦比", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.冶炼强度", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.焦炭负荷", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.干毛焦", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.炼钢铁Si", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.炼钢铁Mn", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.炼钢铁P", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.炼钢铁S", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.铸造铁Si", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.铸造铁Mn", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.铸造铁P", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.铸造铁S", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.炉渣碱度", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.休风情况", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.慢风", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.坐料次数", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.悬料次数", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.崩料次数", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.风口损坏数大", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.风口损坏数中", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.风口损坏数小", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.渣口损坏数大", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.渣口损坏数中", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.渣口损坏数小", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.本厂铸造SI大", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.本厂铸造SI小", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.送炼钢厂SI大", xiangmu, gaolu))
                                                            , 0
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.送炼钢厂SI小", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.折算产量", xiangmu, gaolu))
                                                            , 1
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.工艺称焦比", xiangmu, gaolu))
                                                             , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.S小于002", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.P小于009", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.TI小于055", xiangmu, gaolu))
                                                             , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.PB块", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.纽曼块", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.钛球", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.锰矿", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.硅石", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.白云石", xiangmu, gaolu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.蛇纹石", xiangmu, gaolu))

                                                              // , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.球团矿", xiangmu, gaolu))
                                                              //, (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.国内球团矿", xiangmu, gaolu))
                                                              // , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.进口球团矿", xiangmu, gaolu))
                                                              //  , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.其它块矿", xiangmu, gaolu))
                                                              //   , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.高钛球团矿", xiangmu, gaolu))
                                                              //    , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.高品位钛球", xiangmu, gaolu))
                                                              //     , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.其它熔剂", xiangmu, gaolu))
                                                             , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.萤石", xiangmu, gaolu))
                                                           );


                    SaveData(riqi, gaolu, xiangmu);
                }



            }

            foreach (var xiangmu in xiangmuList)
            {

                if (items.Contains("Q-" + xiangmu))
                    this.技术日报TableAdapter1.Update(
                        (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.合格铁", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.炼钢铁", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.铸造铁", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.号外铁", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.合格率", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.利用系数", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.实物系数", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.原料矿总耗", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.原料矿单耗", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.机烧", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.竖球", xiangmu))
                                                        , 0
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.其它熟料", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.本溪矿", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.其它生料", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.废铁总耗", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.废铁单耗", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.回收率", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.熟料比", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.平均风温", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.炉顶温度", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.热风压力", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.炉顶压力", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.富氧率", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.入炉焦炭总耗", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.入炉焦炭单耗", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.煤粉总耗", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.煤粉单耗", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.焦丁总耗", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.焦丁单耗", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.综合焦炭总耗", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.综合焦炭单耗", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.综合折算焦比", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.冶炼强度", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.焦炭负荷", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.干毛焦", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.炼钢铁Si", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.炼钢铁Mn", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.炼钢铁P", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.炼钢铁S", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.铸造铁Si", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.铸造铁Mn", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.铸造铁P", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.铸造铁S", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.炉渣碱度", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.休风情况", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.慢风", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.坐料次数", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.悬料次数", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.崩料次数", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.风口损坏数大", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.风口损坏数中", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.风口损坏数小", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.渣口损坏数大", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.渣口损坏数中", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.渣口损坏数小", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.本厂铸造SI大", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.本厂铸造SI小", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.送炼钢厂SI大", xiangmu))
                                                        , 0
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.送炼钢厂SI小", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.折算产量", xiangmu))
                                                        , 1
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.工艺称焦比", xiangmu))
                                                         , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.S小于002", xiangmu))
                                                         , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.P小于009", xiangmu))
                                                         , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.TI小于055", xiangmu))
                                                          , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.PB块", xiangmu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.纽曼块", xiangmu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.钛球", xiangmu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.锰矿", xiangmu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.硅石", xiangmu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.白云石", xiangmu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.蛇纹石", xiangmu))

                                                               //, (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.球团矿", xiangmu))
                                                              //, (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.国内球团矿", xiangmu))
                                                              // , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.进口球团矿", xiangmu))
                                                              //  , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.其它块矿", xiangmu))
                                                              //   , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高钛球团矿", xiangmu))
                                                              //    , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高品位钛球", xiangmu))
                                                              //     , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.其它熔剂", xiangmu))

                                                           , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.萤石", xiangmu))
                                                        , riqi, "Q",
                                                          xiangmu);
                else
                    this.技术日报TableAdapter1.Insert(riqi,
                                                         "Q",
                                                         xiangmu
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.合格铁", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.炼钢铁", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.铸造铁", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.号外铁", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.合格率", xiangmu))
                                                        , (decimal?)(decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.利用系数", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.实物系数", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.原料矿总耗", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.原料矿单耗", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.机烧", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.竖球", xiangmu))
                                                        , 0
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.其它熟料", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.本溪矿", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.其它生料", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.废铁总耗", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.废铁单耗", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.回收率", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.熟料比", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.平均风温", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.炉顶温度", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.热风压力", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.炉顶压力", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.富氧率", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.入炉焦炭总耗", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.入炉焦炭单耗", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.煤粉总耗", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.煤粉单耗", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.焦丁总耗", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.焦丁单耗", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.综合焦炭总耗", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.综合焦炭单耗", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.综合折算焦比", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.冶炼强度", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.焦炭负荷", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.干毛焦", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.炼钢铁Si", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.炼钢铁Mn", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.炼钢铁P", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.炼钢铁S", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.铸造铁Si", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.铸造铁Mn", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.铸造铁P", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.铸造铁S", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.炉渣碱度", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.休风情况", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.慢风", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.坐料次数", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.悬料次数", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.崩料次数", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.风口损坏数大", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.风口损坏数中", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.风口损坏数小", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.渣口损坏数大", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.渣口损坏数中", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.渣口损坏数小", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.本厂铸造SI大", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.本厂铸造SI小", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.送炼钢厂SI大", xiangmu))
                                                        , 0
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.送炼钢厂SI小", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.折算产量", xiangmu))
                                                        , 1
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.工艺称焦比", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.S小于002", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.P小于009", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.TI小于055", xiangmu))
                                                        , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.PB块", xiangmu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.纽曼块", xiangmu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.钛球", xiangmu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.锰矿", xiangmu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.硅石", xiangmu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.白云石", xiangmu))
                                                            , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.蛇纹石", xiangmu))

                                                              // , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.球团矿", xiangmu))
                                                              //, (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.国内球团矿", xiangmu))
                                                              // , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.进口球团矿", xiangmu))
                                                              //  , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.其它块矿", xiangmu))
                                                              //   , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高钛球团矿", xiangmu))
                                                              //    , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高品位钛球", xiangmu))
                                                              //     , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.其它熔剂", xiangmu))
                                                              , (decimal?)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.萤石", xiangmu))
                                                       );
                SaveData_Q(riqi, "Q", xiangmu);
            }
            MessageBox.Show("保存完成！");
            this.Cursor = Cursors.Default;                           
        }




         //高炉部分
        public void SaveData(DateTime riqi, int danwei, string xiangmu)
        {
            OracleConnection cn = new OracleConnection();
            cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
            cn.Open();


            try
            {
                OracleCommand updateCmd = new OracleCommand();
                updateCmd.Connection = cn;
                // 机烧矿, 球团矿, 国内球团矿,进口球团矿,PB块,纽曼块,其它块矿,高钛球团矿,高品位钛球,硅石,萤石,蛇纹石,其它熔剂,富氧量,工艺称,煤粉,焦丁,自产湿焦,落地湿焦
                updateCmd.CommandText = "UPDATE 技术日报 SET P77球团矿=:球团矿,P78国内球团矿=:国内球团矿,P79进口球团矿=:进口球团矿,P80其它块矿=:其它块矿,P81高钛球团矿=:高钛球团矿,P82高品位钛球=:高品位钛球,P83其它熔剂=:其它熔剂 WHERE P日期=:日期 and P01单位=:单位 and P02项目=:项目";

               // this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.球团矿", xiangmu, gaolu))
                //(object)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.球团矿", xiangmu)) ?? DBNull.Value;
                updateCmd.Parameters.Add(":球团矿", OracleType.Double, 22).Value = (object)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.球团矿", xiangmu, danwei)) ?? DBNull.Value;
                updateCmd.Parameters.Add(":国内球团矿", OracleType.Double, 22).Value = (object)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.国内球团矿", xiangmu, danwei)) ?? DBNull.Value;
                updateCmd.Parameters.Add(":进口球团矿", OracleType.Double, 22).Value = (object)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.进口球团矿", xiangmu, danwei)) ?? DBNull.Value;
                updateCmd.Parameters.Add(":其它块矿", OracleType.Double, 22).Value = (object)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.其它块矿", xiangmu, danwei)) ?? DBNull.Value;

                updateCmd.Parameters.Add(":高钛球团矿", OracleType.Double, 22).Value = (object)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.高钛球团矿", xiangmu, danwei)) ?? DBNull.Value;
                updateCmd.Parameters.Add(":高品位钛球", OracleType.Double, 22).Value = (object)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.高品位钛球", xiangmu, danwei)) ?? DBNull.Value;
                updateCmd.Parameters.Add(":其它熔剂", OracleType.Double, 22).Value = (object)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.其它熔剂", xiangmu, danwei)) ?? DBNull.Value;
               
                updateCmd.Parameters.Add(":日期", OracleType.DateTime).Value = riqi;
                updateCmd.Parameters.Add(":单位", OracleType.VarChar).Value = danwei + "#"; 
                updateCmd.Parameters.Add(":项目", OracleType.VarChar).Value = xiangmu;
                updateCmd.ExecuteNonQuery();

            }
            catch
            {

                return;
            }

            cn.Close();
        }

        // 全厂部分
        public void SaveData_Q(DateTime riqi, string  danwei, string xiangmu)
        {
            OracleConnection cn = new OracleConnection();
            cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
            cn.Open();


            try
            {
                OracleCommand updateCmd = new OracleCommand();
                updateCmd.Connection = cn;
                // 机烧矿, 球团矿, 国内球团矿,进口球团矿,PB块,纽曼块,其它块矿,高钛球团矿,高品位钛球,硅石,萤石,蛇纹石,其它熔剂,富氧量,工艺称,煤粉,焦丁,自产湿焦,落地湿焦
                updateCmd.CommandText = "UPDATE 技术日报 SET P77球团矿=:球团矿,P78国内球团矿=:国内球团矿,P79进口球团矿=:进口球团矿,P80其它块矿=:其它块矿,P81高钛球团矿=:高钛球团矿,P82高品位钛球=:高品位钛球,P83其它熔剂=:其它熔剂 WHERE P日期=:日期 and P01单位=:单位 and P02项目=:项目";

                // this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高炉{1}.球团矿", xiangmu, gaolu))
                //(object)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.球团矿", xiangmu)) ?? DBNull.Value;
                updateCmd.Parameters.Add(":球团矿", OracleType.Double, 22).Value = (object)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.球团矿", xiangmu)) ?? DBNull.Value;
                updateCmd.Parameters.Add(":国内球团矿", OracleType.Double, 22).Value = (object)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.国内球团矿", xiangmu)) ?? DBNull.Value;
                updateCmd.Parameters.Add(":进口球团矿", OracleType.Double, 22).Value = (object)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.进口球团矿", xiangmu)) ?? DBNull.Value;
                updateCmd.Parameters.Add(":其它块矿", OracleType.Double, 22).Value = (object)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.其它块矿", xiangmu)) ?? DBNull.Value;

                updateCmd.Parameters.Add(":高钛球团矿", OracleType.Double, 22).Value = (object)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高钛球团矿", xiangmu)) ?? DBNull.Value;
                updateCmd.Parameters.Add(":高品位钛球", OracleType.Double, 22).Value = (object)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.高品位钛球", xiangmu)) ?? DBNull.Value;
                updateCmd.Parameters.Add(":其它熔剂", OracleType.Double, 22).Value = (object)this.JiShuRiBao.myModel.GetNullTagValue(string.Format("{0}.其它熔剂", xiangmu)) ?? DBNull.Value;

                updateCmd.Parameters.Add(":日期", OracleType.DateTime).Value = riqi;
                updateCmd.Parameters.Add(":单位", OracleType.VarChar).Value = danwei;
                updateCmd.Parameters.Add(":项目", OracleType.VarChar).Value = xiangmu;
                updateCmd.ExecuteNonQuery();

            }
            catch
            {

                return;
            }

            cn.Close();
        }

        private 技术日报内容 dataSource = null;
        private string[] xiufeng = new string[9];
        private string[] manfeng = new string[9];
        private string[] qitaqingkuang = new string[2];

        public void setDataSource(技术日报内容 n)
        {
            //初始化
            xiufeng.Initialize();
            manfeng.Initialize();
            qitaqingkuang.Initialize();

            XiaohaoGaolu.DataSetXh1 xmf = new LTZN.XiaohaoGaolu.DataSetXh1();
            XiaohaoGaolu.DataSetXh1TableAdapters.休风TableAdapter xadp = new LTZN.XiaohaoGaolu.DataSetXh1TableAdapters.休风TableAdapter();
            XiaohaoGaolu.DataSetXh1TableAdapters.慢风TableAdapter madp = new LTZN.XiaohaoGaolu.DataSetXh1TableAdapters.慢风TableAdapter();
            //    XiaohaoGaolu.DataSetXh1TableAdapters.全厂日消耗TableAdapter qadp = new LTZN.XiaohaoGaolu.DataSetXh1TableAdapters.全厂日消耗TableAdapter();

            xadp.FillByRq(xmf.休风, n.riqi);
            madp.FillByRq(xmf.慢风, n.riqi);
            //  qadp.FillByRq(xmf.全厂日消耗, n.riqi);


            foreach (技术日报内容项 x in n)
            {
                x.Convert0ToNull();
            }
            this.技术日报BindingSource.DataSource = n;
            dataSource = n;

            List<ReportParameter> para = new List<ReportParameter>();

            para.Add(new ReportParameter("Riqi", n.riqi.ToString("yyyy年MM月dd日")));

            int i = 1;

            string strxiufeng = "";
            foreach (XiaohaoGaolu.DataSetXh1.休风Row r in xmf.休风)
            {
               
                if (i <= 9)
                {
                    strxiufeng += i + "、" + r.高炉 + "炉" + r.时间.ToString("H:mm") + "分休风　" + r.间隔.ToString() + "分钟," + (r.Is原因Null() ? "" : r.原因) + "  ";
                    i++;
                }
            }

            para.Add(new ReportParameter("xiufeng1", strxiufeng));

            //while (i <= 9)
            //{
            //    para.Add(new ReportParameter("xiufeng" + i, ""));
            //    xiufeng[i - 1] = "";  
            //    i++;
            //}

            i = 1;
            string strmanfeng = "";
            foreach (XiaohaoGaolu.DataSetXh1.慢风Row r in xmf.慢风)
            {
                
                if (i <= 9)
                {            
                    strmanfeng += i + "、" + r.高炉 + "炉" + r.时间.ToString("H:mm") + "分慢风　" + r.间隔.ToString() + "小时," + (r.Is原因Null() ? "" : r.原因) + "  ";  
                    i++;
                }
            }

            para.Add(new ReportParameter("manfeng1", strmanfeng));
            //while (i <= 9)
            //{
            //    para.Add(new ReportParameter("manfeng" + i, ""));
            //    manfeng[i - 1] = "";   
            //    i++;
            //}

            OracleConnection cn = new OracleConnection();
            cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
            cn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = cn;
            cmd.CommandText = "select P09备注1,P10备注2 from 全厂日消耗 where P01日期=:rq";
            cmd.Parameters.Add(":rq", OracleType.DateTime).Value = n.riqi;
            OracleDataReader dr = cmd.ExecuteReader();
            if (dr.Read())
            {
                if (!dr.IsDBNull(0) && dr.GetString(0) != "")
                {
                    string qtStr = dr.GetString(0);
                    para.Add(new ReportParameter("qita1", qtStr));
                    qitaqingkuang[0] = qtStr;
                }
                else
                {
                    para.Add(new ReportParameter("qita1", ""));
                    qitaqingkuang[0] = "";

                }

                if (!dr.IsDBNull(1) && dr.GetString(1) != "")
                {
                    string qtStr = dr.GetString(1);
                    para.Add(new ReportParameter("qita2", qtStr));
                    qitaqingkuang[1] = qtStr;
                }
                else
                {
                    para.Add(new ReportParameter("qita2", ""));
                    qitaqingkuang[1] = "";
                }
            }
            dr.Close();
            cn.Close();

            String leader = LTZN.Data.SysParameter.GetStrParameter("单位负责人", "陈学清");
            String dayReportAuthor = LTZN.Data.SysParameter.GetStrParameter("日报负责人","刘红芬");
            String monthReportAuthor = LTZN.Data.SysParameter.GetStrParameter("月报负责人", "王文英");


            string yejiao = "单位负责人：“" + leader + "”　 制表人：“" + dayReportAuthor + "”　 第 ＋　＋ 页，共一页";
            para.Add(new ReportParameter("yejiao", yejiao));
            para.Add(new ReportParameter("actUser", dayReportAuthor));
            this.reportViewer1.LocalReport.SetParameters(para);
            this.reportViewer1.RefreshReport();
            this.reportViewer1.Refresh();

        }

        //private void btnOldRiBao_Click(object sender, EventArgs e)
        //{
        //    炼铁智能主窗体 mainFrm = this.ParentForm as 炼铁智能主窗体;
        //    if(mainFrm!=null)
        //        mainFrm.ShowForm(typeof(数据修改窗体), "数据修改窗体");
        //}

        private void mnuToExcel_Click(object sender, EventArgs e)
        {
            if (dataSource == null)
                throw new Exception("没有数据！");

            string title = "炼铁技术日报" + dataSource.riqi.ToString("yyyy年MM月dd日");
            string fn = title + ".xls";
            // string fullfn = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + fn;

            saveFileDialog1.FileName = fn;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {

                this.Cursor = Cursors.WaitCursor;
                Stream file = saveFileDialog1.OpenFile();
                file.Write(Properties.Resources.炼铁技术日报, 0, Properties.Resources.炼铁技术日报.Length);
                file.Close();

                object m_objOpt = System.Reflection.Missing.Value;

                ApplicationClass app = new ApplicationClass();
                app.Visible = false;
                Workbook wb = app.Workbooks.Open(saveFileDialog1.FileName, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt);
                wb.Title = title;
                Worksheet ws = (Worksheet)wb.Sheets[1];

                ws.Cells.set_Item(3, "AG", dataSource.riqi.ToString("yyyy年MM月dd日"));

                int row = 6;
                foreach (技术日报内容项 x in dataSource)
                {
                    ws.Cells.set_Item(row, "B", x.P01单位);
                    ws.Cells.set_Item(row, "E", x.P02项目);
                    ws.Cells.set_Item(row, "I", x.P03合格铁);
                    ws.Cells.set_Item(row, "Q", x.P04炼钢铁);
                    ws.Cells.set_Item(row, "T", x.P05铸造铁);
                    ws.Cells.set_Item(row, "Y", x.P06号外铁);
                    ws.Cells.set_Item(row, "AB", x.P07合格率);
                    ws.Cells.set_Item(row, "AK", x.P08高炉利用系数);
                    ws.Cells.set_Item(row, "AM", x.P09高炉实物系数);
                    ws.Cells.set_Item(row, "AQ", x.P10原料矿合计总耗);
                    ws.Cells.set_Item(row, "AT", x.P11原料矿合计单耗);
                    ws.Cells.set_Item(row, "AW", x.P12原料矿机烧);
                    ws.Cells.set_Item(row, "BA", x.P13原料矿竖炉球);
                    ws.Cells.set_Item(row, "BD", x.P15原料矿其它熟料);
                    ws.Cells.set_Item(row, "BF", x.P69PB块);
                    ws.Cells.set_Item(row, "BI", x.P70纽曼块);
                    ws.Cells.set_Item(row, "BL", x.P71钛球);
                    ws.Cells.set_Item(row, "BP", x.P72锰矿);
                    ws.Cells.set_Item(row, "BU", x.P20回收率);
                    ws.Cells.set_Item(row, "CB", x.P21熟料比);
                    ws.Cells.set_Item(row, "CE", x.P22平均风温);
                    ws.Cells.set_Item(row, "CK", x.P23炉顶温度);
                    ws.Cells.set_Item(row, "CP", x.P24热风压力);
                    ws.Cells.set_Item(row, "CT", x.P25炉顶压力);
                    ws.Cells.set_Item(row, "CX", x.P26富氧率);
                    ws.Cells.set_Item(row, "DB", x.P27入炉焦炭总耗);
                    ws.Cells.set_Item(row, "DD", x.P28入炉焦炭单耗);
                    ws.Cells.set_Item(row, "DJ", x.P29煤粉总耗);
                    ws.Cells.set_Item(row, "DL", x.P30煤粉单耗);

                    ws.Cells.set_Item(row + 13, "B", x.P01单位);
                    ws.Cells.set_Item(row + 13, "F", x.P02项目);
                    ws.Cells.set_Item(row + 13, "J", x.P31焦丁总耗);
                    ws.Cells.set_Item(row + 13, "L", x.P32焦丁单耗);
                    ws.Cells.set_Item(row + 13, "O", x.P33综合焦炭总耗);
                    ws.Cells.set_Item(row + 13, "R", x.P34综合焦炭单耗);
                    ws.Cells.set_Item(row + 13, "V", x.P35综合折算焦比);
                    ws.Cells.set_Item(row + 13, "AA", x.P36冶炼强度);
                    ws.Cells.set_Item(row + 13, "AE", x.P37焦炭负荷);
                    ws.Cells.set_Item(row + 13, "AI", x.P38干毛焦);
                    ws.Cells.set_Item(row + 13, "AO", x.P39炼钢铁SI);
                    ws.Cells.set_Item(row + 13, "AR", x.P40炼钢铁MN);
                    ws.Cells.set_Item(row + 13, "AS", x.P41炼钢铁P);
                    ws.Cells.set_Item(row + 13, "AV", x.P42炼钢铁S);
                    ws.Cells.set_Item(row + 13, "AZ", x.P73硅石);
                    ws.Cells.set_Item(row + 13, "BB", x.P74白云石);
                    ws.Cells.set_Item(row + 13, "BC", x.P75蛇纹石);
                  //  ws.Cells.set_Item(row + 13, "BE", x.P46铸造铁S);
                    ws.Cells.set_Item(row + 13, "BH", x.P47炉渣碱度);
                    ws.Cells.set_Item(row + 13, "BK", x.P48休风情况);
                    ws.Cells.set_Item(row + 13, "BO", x.P49慢风);
                    ws.Cells.set_Item(row + 13, "BR", x.P50坐料次数);
                    ws.Cells.set_Item(row + 13, "BW", x.P51悬料次数);
                    ws.Cells.set_Item(row + 13, "BZ", x.P52崩料次数);
                    ws.Cells.set_Item(row + 13, "CD", x.P53风口损坏数大);
                    ws.Cells.set_Item(row + 13, "CG", x.P54风口损坏数中);
                    ws.Cells.set_Item(row + 13, "CH", x.P55风口损坏数小);
                    ws.Cells.set_Item(row + 13, "CN", x.P56渣口损坏数大);
                    ws.Cells.set_Item(row + 13, "CO", x.P57渣口损坏数中);
                    ws.Cells.set_Item(row + 13, "CR", x.P58渣口损坏数小);
                    ws.Cells.set_Item(row + 13, "CV", x.P65工艺称焦比);

                    double b1 = x.P59本厂铸造SI大 == null ? 0 : x.P59本厂铸造SI大.Value;
                    double b2 = x.P60本厂铸造SI小 == null ? 0 : x.P60本厂铸造SI小.Value;
                    double? b3 = b1 + b2;
                    if (b3 == 0)
                        b3 = null;

                    double l1 = x.P61送炼钢厂SI大 == null ? 0 : x.P61送炼钢厂SI大.Value;
                    double l2 = x.P62送炼钢厂SI中 == null ? 0 : x.P62送炼钢厂SI中.Value;
                    double l3 = x.P63送炼钢厂SI小 == null ? 0 : x.P63送炼钢厂SI小.Value;
                    double? l4 = l1 + l2 + l3;
                    if (l4 == 0)
                        l4 = null;

                    ws.Cells.set_Item(row + 13, "DF", b3);
                    ws.Cells.set_Item(row + 13, "DK", l4);
                    row++;
                }

                ws.Cells.set_Item(30, "M", xiufeng[0]);
                ws.Cells.set_Item(31, "M", xiufeng[1]);
                ws.Cells.set_Item(32, "M", xiufeng[2]);
                ws.Cells.set_Item(30, "AZ", xiufeng[3]);
                ws.Cells.set_Item(31, "AZ", xiufeng[4]);
                ws.Cells.set_Item(32, "AZ", xiufeng[5]);
                ws.Cells.set_Item(30, "CJ", xiufeng[6]);
                ws.Cells.set_Item(31, "CJ", xiufeng[7]);
                ws.Cells.set_Item(32, "CJ", xiufeng[8]);

                ws.Cells.set_Item(33, "M", manfeng[0]);
                ws.Cells.set_Item(34, "M", manfeng[1]);
                ws.Cells.set_Item(35, "M", manfeng[2]);
                ws.Cells.set_Item(33, "AZ", manfeng[3]);
                ws.Cells.set_Item(34, "AZ", manfeng[4]);
                ws.Cells.set_Item(35, "AZ", manfeng[5]);
                ws.Cells.set_Item(33, "CJ", manfeng[6]);
                ws.Cells.set_Item(34, "CJ", manfeng[7]);
                ws.Cells.set_Item(35, "CJ", manfeng[8]);

                ws.Cells.set_Item(36, "M", qitaqingkuang[0]);
                ws.Cells.set_Item(37, "M", qitaqingkuang[1]);

                String leader = LTZN.Data.SysParameter.GetStrParameter("单位负责人", "马保良");
                String dayReportAuthor = LTZN.Data.SysParameter.GetStrParameter("日报负责人", "刘红芬");
                String monthReportAuthor = LTZN.Data.SysParameter.GetStrParameter("月报负责人", "刘红芬");

                string yejiao = "单位负责人：" + leader + "“　“制表人： " + dayReportAuthor + "”　”第“＋　＋”页，共一页”";
                ws.Cells.set_Item(38, "B", yejiao);

                wb.Save();
                wb.Saved = true;
                app.Visible = true;
                this.Cursor = Cursors.Default;
            }
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }
    }
}
