using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Reporting.WinForms;
using System.Data.OracleClient;
using LTZN.Repository;
using Microsoft.Office.Interop.Excel;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace LTZN.技术月报
{
    public partial class MonthReportForm : Form
    {
        public MonthReportForm()
        {
            InitializeComponent();
            LtznUserManager.instance.RegisterHandler(this, UserChange);
        }

        //用户变动时
        void UserChange(LtznUser ltznUser)
        {
            if (ltznUser != null && ltznUser.IsInRole("统计"))
            {
                btnCreate.Enabled = true;
                btnSave.Enabled = true;
                btnOldRiBao.Enabled = true;
            }
            else
            {
                btnCreate.Enabled = false;
                btnSave.Enabled = false;
                btnOldRiBao.Enabled = false;
            }
            if (ltznUser != null && ltznUser.IsInRole("技术月报模型"))
            {
                btnModel.Enabled = true;
            }
            else
            {
                btnModel.Enabled = false;
            }
        
        }

        private 技术月报 _jiShuYueBao = null;
        public 技术月报 JiShuYueBao
        {
            get
            {
                if (_jiShuYueBao == null)
                {
                    _jiShuYueBao = new 技术月报();
                }
                return _jiShuYueBao;
            }
        }
        
        private void btnShowReport_Click(object sender, EventArgs e)
        {
           JiShuYueBaoTable table = new JiShuYueBaoTable();
           table.LoadByMonth(Convert.ToInt32(TextBox_nian.Text), Convert.ToInt32(TextBox_yue.Text));
           this.技术月报内容BindingSource.DataSource = table;
           List<ReportParameter> para = new List<ReportParameter>();
           para.Add(new ReportParameter("actUser", LTZN.Data.SysParameter.GetStrParameter("日报负责人")));
             para.Add(new ReportParameter("actLeader", LTZN.Data.SysParameter.GetStrParameter("单位负责人")));
           try
           {
               this.reportViewer1.LocalReport.SetParameters(para);
           }
           catch (Exception ex)
           {
               Console.Write(ex);
           }
           reportViewer1.RefreshReport();
        }

        private void btnOldRiBao_Click(object sender, EventArgs e)
        {
            炼铁智能主窗体 mainFrm = this.ParentForm as 炼铁智能主窗体;
            if (mainFrm != null)
                mainFrm.ShowForm(typeof(Form1), "技术月报");
        }

        /// <summary>
        /// 生成月报
        /// </summary>
        /// 
        string a = "";
        private void btnCreate_Click(object sender, EventArgs e)
        {
            try
            {
                a = "1";
                this.Cursor = Cursors.WaitCursor;
                int nian = Convert.ToInt32(this.TextBox_nian.Text);
                int yue = Convert.ToInt32(this.TextBox_yue.Text);
                this.JiShuYueBao.myModel.Init();
                a = "2";
                this.JiShuYueBao.getDataBy(nian, yue);
                a = "3";
                this.JiShuYueBao.ModelCalc();
                a = "4";
                JiShuYueBaoTable table = ModelToTable(nian, yue);
                a = "5";
                this.Cursor = Cursors.Default;
            }
            catch (Exception ex)
            {
                MessageBox.Show("生成月报出错" + a + ex.ToString());
                return;
            }
          
        }

        private void btnModel_Click(object sender, EventArgs e)
        {
            if (this.JiShuYueBao.myModel != null)
            {
                CalModelForm form = new CalModelForm(this.JiShuYueBao.myModel);
                form.ShowDialog();
                int nian = Convert.ToInt32(this.TextBox_nian.Text);
                int yue = Convert.ToInt32(this.TextBox_yue.Text);
                JiShuYueBaoTable table = ModelToTable(nian, yue);
            }
        }

        private void MonthReportForm_Load(object sender, EventArgs e)
        {
            DateTime shangyue=DateTime.Now.AddMonths(-1);
            this.TextBox_nian.Text =shangyue.Year.ToString();
            this.TextBox_yue.Text = shangyue.Month.ToString();
        }


        private void btnSave_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            int nian = Convert.ToInt32(this.TextBox_nian.Text);
            int yue = Convert.ToInt32(this.TextBox_yue.Text);
            JiShuYueBaoTable table = ModelToTable(nian, yue);
            table.Save();
            MessageBox.Show("保存完成！");
            this.Cursor = Cursors.Default;
        }

        private JiShuYueBaoTable ModelToTable(int nian, int yue)
        {
            JiShuYueBaoTable table = new JiShuYueBaoTable();
            List<string> xiangmuList = new List<string>() { "本月", "累计" };
            for (int gaolu = 1; gaolu < 8; gaolu++)
            {
                if (gaolu == 2 || gaolu == 4) continue;
                foreach (var xiangmu in xiangmuList)
                {
                    string danwei = string.Format("{0}#", gaolu);
                    string zuzhi = string.Format("高炉{0}.", gaolu);
                    if (gaolu == 7)
                    {
                        danwei = "Q";
                        zuzhi = "";
                    }
                    JiShuYueBao yuebao = table.GetJishuYueBao(nian, yue, danwei, xiangmu);
                    yuebao.P03全铁产量 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}全铁产量", xiangmu, zuzhi));
                    yuebao.P04合格铁 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}合格铁", xiangmu, zuzhi));
                    yuebao.P05制钢铁 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}制钢铁", xiangmu, zuzhi));
                    yuebao.P06铸造铁 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}铸造铁", xiangmu, zuzhi));
                    yuebao.P07折合产量 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}折算产量", xiangmu, zuzhi));
                    yuebao.P08合格率 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}合格率", xiangmu, zuzhi));
                    yuebao.P09一级品率 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}一级品率", xiangmu, zuzhi));
                    yuebao.P10送炼钢优质铁水率 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}送炼钢优质铁水率", xiangmu, zuzhi));
                    yuebao.P11高炉利用系数 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}利用系数", xiangmu, zuzhi));
                    yuebao.P12入炉焦炭冶炼强度 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}入炉焦炭冶炼强度", xiangmu, zuzhi));
                    yuebao.P13综合焦炭冶炼强度 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}综合焦炭冶炼强度", xiangmu, zuzhi));
                    yuebao.P14折算综合焦比 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}折算综合焦比", xiangmu, zuzhi));
                    yuebao.P15折合入炉焦比 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}折算入炉焦比", xiangmu, zuzhi));
                    yuebao.P16入炉矿品位 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}入炉矿品位", xiangmu, zuzhi));
                    yuebao.P17熟料比 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}熟料比", xiangmu, zuzhi));
                    yuebao.P18入炉焦炭负荷 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}入炉焦炭负荷", xiangmu, zuzhi));
                    yuebao.P19综合焦炭负荷 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}综合焦炭负荷", xiangmu, zuzhi));
                    yuebao.P20休风时间 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}休风时间", xiangmu, zuzhi));
                    yuebao.P21休风率 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}休风率", xiangmu, zuzhi));
                    yuebao.P22慢风时间 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}慢风时间", xiangmu, zuzhi));
                    yuebao.P23慢风率 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}慢风率", xiangmu, zuzhi));
                    yuebao.P24煤气成分CO2 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}CO2", xiangmu, zuzhi));
                    yuebao.P25煤气成分计算1 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}CO2率", xiangmu, zuzhi));
                    yuebao.P26生铁表面质量 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}生铁表面质量", xiangmu, zuzhi));
                    yuebao.P27深料线作业率 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}深料线作业率", xiangmu, zuzhi));
                    yuebao.P28回收率 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}回收率", xiangmu, zuzhi));
                    yuebao.P29冷风流量 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}冷风流量", xiangmu, zuzhi));
                    yuebao.P30平均风温 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}平均风温", xiangmu, zuzhi));
                    yuebao.P31风速 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}风速", xiangmu, zuzhi));
                    yuebao.P32热风压力 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}热风压力", xiangmu, zuzhi));
                    yuebao.P33炉顶压力 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}炉顶压力", xiangmu, zuzhi));
                    yuebao.P34炉顶温度 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}炉顶温度", xiangmu, zuzhi));
                    yuebao.P35富氧率 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}富氧率", xiangmu, zuzhi));
                    yuebao.P36高压率 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}高压率", xiangmu, zuzhi));
                    yuebao.P37悬料次数 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}悬料次数", xiangmu, zuzhi));
                    yuebao.P38坐料次数 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}坐料次数", xiangmu, zuzhi));
                    yuebao.P39崩料次数 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}崩料次数", xiangmu, zuzhi));
                    yuebao.P40风口大套损坏数 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}风口损坏数大", xiangmu, zuzhi));
                    yuebao.P41风口中套损坏数 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}风口损坏数中", xiangmu, zuzhi));
                    yuebao.P42风口小套损坏数 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}风口损坏数小", xiangmu, zuzhi));
                    yuebao.P43渣口大套损坏数 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}渣口损坏数大", xiangmu, zuzhi));
                    yuebao.P44渣口中套损坏数 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}渣口损坏数中", xiangmu, zuzhi));
                    yuebao.P45渣口小套损坏数 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}渣口损坏数小", xiangmu, zuzhi));
                    yuebao.P46渣铁比 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}渣铁比", xiangmu, zuzhi));
                    yuebao.P47灰铁比 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}灰铁比", xiangmu, zuzhi));
                    yuebao.P48原料总耗 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}原料总耗", xiangmu, zuzhi));
                    yuebao.P49原料单耗 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}原料单耗", xiangmu, zuzhi));
                    yuebao.P50机烧消耗 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}机烧", xiangmu, zuzhi));
                    yuebao.P51竖炉球消耗 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}竖球", xiangmu, zuzhi));
                    yuebao.P52印度球消耗 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}印度球", xiangmu, zuzhi));
                    yuebao.P53其它熟料消耗 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}其它熟料", xiangmu, zuzhi));
                    yuebao.P54本溪矿消耗 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}本溪矿", xiangmu, zuzhi));
                    yuebao.P55其它生料消耗 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}其它生料", xiangmu, zuzhi));
                    yuebao.P56废铁总耗 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}废铁", xiangmu, zuzhi));
                    yuebao.P57废铁单耗 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}废铁单耗", xiangmu, zuzhi));
                    yuebao.P58金属收得率 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}金属收得率", xiangmu, zuzhi));
                    yuebao.P59综合焦炭总耗 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}综合焦炭", xiangmu, zuzhi));
                    yuebao.P60七号称 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}七号称", xiangmu, zuzhi));
                    yuebao.P61干毛焦 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}干毛焦", xiangmu, zuzhi));
                    yuebao.P62干焦粉 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}干焦粉", xiangmu, zuzhi));
                    yuebao.P63入炉焦炭总耗 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}入炉焦炭", xiangmu, zuzhi));
                    yuebao.P64入炉焦炭单耗 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}入炉焦炭单耗", xiangmu, zuzhi));
                    yuebao.P65煤粉总耗 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}煤粉", xiangmu, zuzhi));
                    yuebao.P66焦丁总耗 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}焦丁", xiangmu, zuzhi));
                    yuebao.P67焦丁单耗 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}焦丁单耗", xiangmu, zuzhi));
                    yuebao.P68燃料比 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}燃料比", xiangmu, zuzhi));
                    yuebao.P69铁成分SI = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}Fe_Si", xiangmu, zuzhi));
                    yuebao.P70铁成分MN = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}Fe_Mn", xiangmu, zuzhi));
                    yuebao.P71铁成分P = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}Fe_P", xiangmu, zuzhi));
                    yuebao.P72铁成分S = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}Fe_S", xiangmu, zuzhi));
                    yuebao.P73渣成分CAO = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}Zha_CaO", xiangmu, zuzhi));
                    yuebao.P74渣成分SIO2 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}Zha_SiO2", xiangmu, zuzhi));
                    yuebao.P75渣成分AL2O3 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}Zha_Al2O3", xiangmu, zuzhi));
                    yuebao.P76渣成分MGO = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}Zha_MgO", xiangmu, zuzhi));
                    yuebao.P77渣成分FEO = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}Zha_FeO", xiangmu, zuzhi));
                    yuebao.P78渣成分S = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}Zha_S", xiangmu, zuzhi));
                    yuebao.P79含SI偏差CAO = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}碱度", xiangmu, zuzhi));
                    yuebao.P80含SI偏差制铁量 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}制钢铁Si偏差", xiangmu, zuzhi));
                    yuebao.P81含SI偏差铸造铁 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}铸造铁Si偏差", xiangmu, zuzhi));
                    yuebao.P82工艺称焦比 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}工艺称焦比", xiangmu, zuzhi));
                    yuebao.富氧量 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}富氧量", xiangmu, zuzhi));
                    yuebao.高炉实际容积 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}高炉实际容积", xiangmu, zuzhi));
                    yuebao.高炉有效容积 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}高炉有效容积", xiangmu, zuzhi));
                    yuebao.高压操作时间 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}高压操作时间", xiangmu, zuzhi));
                    yuebao.理论渣量 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}理论渣量", xiangmu, zuzhi));
                    yuebao.煤粉单耗 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}煤粉单耗 ", xiangmu, zuzhi));
                    yuebao.全部料线 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}全部料线", xiangmu, zuzhi));
                    yuebao.入炉矿含铁 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}入炉矿含铁", xiangmu, zuzhi));
                    yuebao.深料线 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}深料线", xiangmu, zuzhi));
                    yuebao.收入含铁 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}收入含铁", xiangmu, zuzhi));
                    yuebao.瓦斯灰 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}瓦斯灰", xiangmu, zuzhi));
                    yuebao.一级铁 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}一级铁", xiangmu, zuzhi));
                    yuebao.优质铁 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}优质铁", xiangmu, zuzhi));
                    yuebao.有效工作时间 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}有效工作时间", xiangmu, zuzhi));
                    yuebao.支出含铁 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}支出含铁", xiangmu, zuzhi));
                    yuebao.综合焦炭单耗 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}综合焦炭单耗", xiangmu, zuzhi));
                    yuebao.优质铁 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}优质铁", xiangmu, zuzhi));
                    yuebao.PB块 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}PB块", xiangmu, zuzhi));
                    yuebao.纽曼块 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}纽曼块", xiangmu, zuzhi));
                    yuebao.锰矿 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}锰矿", xiangmu, zuzhi));
                    yuebao.钛球 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}钛球", xiangmu, zuzhi));
                    yuebao.溶剂总耗 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}溶剂总耗", xiangmu, zuzhi));
                    yuebao.溶剂单耗 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}溶剂单耗", xiangmu, zuzhi));
                    yuebao.硅石 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}硅石", xiangmu, zuzhi));
                    yuebao.白云石 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}白云石", xiangmu, zuzhi));
                    yuebao.蛇纹石 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}蛇纹石", xiangmu, zuzhi));
                    //2016-11-24 添加   P76萤石,P77球团矿,P78国内球团矿,P79进口球团矿,P80其它块矿,P81高钛球团矿,P82高品位钛球,P83其它熔剂
                    yuebao.萤石 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}萤石", xiangmu, zuzhi));
                    yuebao.球团矿 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}球团矿", xiangmu, zuzhi));
                    yuebao.国内球团矿 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}国内球团矿", xiangmu, zuzhi));
                    yuebao.进口球团矿 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}进口球团矿", xiangmu, zuzhi));
                    yuebao.其它块矿 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}其它块矿", xiangmu, zuzhi));
                    yuebao.高钛球团矿 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}高钛球团矿", xiangmu, zuzhi));
                    yuebao.高品位钛球 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}高品位钛球", xiangmu, zuzhi));
                    yuebao.其它熔剂 = this.JiShuYueBao.myModel.GetNullTagValue(string.Format("{0}.{1}其它熔剂", xiangmu, zuzhi));
                }
            }

            this.技术月报内容BindingSource.DataSource = table;
            String leader = LTZN.Data.SysParameter.GetStrParameter("单位负责人", "陈学清");
            String dayReportAuthor = LTZN.Data.SysParameter.GetStrParameter("日报负责人", "刘红芬");
            String monthReportAuthor = LTZN.Data.SysParameter.GetStrParameter("月报负责人", "刘彦丽");

            List<ReportParameter> para = new List<ReportParameter>();
            para.Add(new ReportParameter("actUser", monthReportAuthor));
            para.Add(new ReportParameter("actLeader", leader));
            try
            {
                this.reportViewer1.LocalReport.SetParameters(para);
            }
            catch (Exception ex)
            {
                Console.Write(ex);
            }
            reportViewer1.RefreshReport();

            return table;
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            if (reportViewer1.Visible)
                reportViewer1.PrintDialog();
            else
                webBrowser1.ShowPrintPreviewDialog();
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            
            SaveToExcel();
        }
        private void old_excel()
        {
            JiShuYueBaoTable table = this.技术月报内容BindingSource.DataSource as JiShuYueBaoTable;
            if (table != null)
            {
                int nian = Convert.ToInt32(this.TextBox_nian.Text);
                int yue = Convert.ToInt32(this.TextBox_yue.Text);
                string title = string.Format("炼铁技术月报{0}年{1}月", nian, yue);
                string fn = title + ".xls";
                // string fullfn = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + fn;

                saveFileDialog1.FileName = fn;
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {

                    this.Cursor = Cursors.WaitCursor;
                    Stream file = saveFileDialog1.OpenFile();
                    file.Write(Properties.Resources.炼铁技术月报, 0, Properties.Resources.炼铁技术月报.Length);
                    file.Close();

                    object m_objOpt = System.Reflection.Missing.Value;

                    ApplicationClass app = new ApplicationClass();
                    app.Visible = false;
                    Workbook wb = app.Workbooks.Open(saveFileDialog1.FileName, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt);
                    wb.Title = title;
                    Worksheet ws = (Worksheet)wb.Sheets[1];

                    ws.Cells[3, 7] = nian + "年" + yue + "月";
 
                    String monthReportAuthor = LTZN.Data.SysParameter.GetStrParameter("月报负责人", "刘彦丽");
                    ws.Cells[3, 27] = monthReportAuthor;
                    int row = 6;
                    int rowgroup2 = 12;
                    int rowgroup3 = 24;
                    foreach (var item in table)
                    {
                        
                        ws.Cells.set_Item(row, "D", item.P03全铁产量);
                        ws.Cells.set_Item(row, "E", item.P05制钢铁);
                        ws.Cells.set_Item(row, "F", item.P06铸造铁);
                        ws.Cells.set_Item(row, "G", item.P07折合产量);
                        ws.Cells.set_Item(row, "H", item.P08合格率);
                        ws.Cells.set_Item(row, "I", item.P09一级品率);
                        ws.Cells.set_Item(row, "J", item.P10送炼钢优质铁水率);
                        ws.Cells.set_Item(row, "K", item.P11高炉利用系数);
                        ws.Cells.set_Item(row, "L", item.P12入炉焦炭冶炼强度);
                        ws.Cells.set_Item(row, "M", item.P13综合焦炭冶炼强度);
                        ws.Cells.set_Item(row, "N", item.P14折算综合焦比);
                        ws.Cells.set_Item(row, "O", item.P15折合入炉焦比);
                        ws.Cells.set_Item(row, "P", item.P16入炉矿品位);
                        ws.Cells.set_Item(row, "Q", item.P17熟料比);
                        ws.Cells.set_Item(row, "R", item.P18入炉焦炭负荷);
                        ws.Cells.set_Item(row, "S", item.P19综合焦炭负荷);
                        ws.Cells.set_Item(row, "T", item.P20休风时间);
                        ws.Cells.set_Item(row, "U", item.P21休风率);
                        ws.Cells.set_Item(row, "V", item.P22慢风时间);
                        ws.Cells.set_Item(row, "W", item.P23慢风率);
                        ws.Cells.set_Item(row, "X", item.P24煤气成分CO2);
                        ws.Cells.set_Item(row, "Y", item.P25煤气成分计算1);
                        ws.Cells.set_Item(row, "Z", item.P27深料线作业率);
                        ws.Cells.set_Item(row, "AA", item.P28回收率);
                        ws.Cells.set_Item(row, "AB", item.P29冷风流量);
                        ws.Cells.set_Item(row, "AC", item.P30平均风温);

                        ws.Cells.set_Item(row + rowgroup2, "D", item.P32热风压力);
                        ws.Cells.set_Item(row + rowgroup2, "E", item.P33炉顶压力);
                        ws.Cells.set_Item(row + rowgroup2, "F", item.P34炉顶温度);
                        ws.Cells.set_Item(row + rowgroup2, "G", item.P35富氧率);
                        ws.Cells.set_Item(row + rowgroup2, "H", item.P36高压率);
                        ws.Cells.set_Item(row + rowgroup2, "I", item.P37悬料次数);
                        ws.Cells.set_Item(row + rowgroup2, "J", item.P38坐料次数);
                        ws.Cells.set_Item(row + rowgroup2, "K", item.P39崩料次数);
                        ws.Cells.set_Item(row + rowgroup2, "L", item.P40风口大套损坏数);
                        ws.Cells.set_Item(row + rowgroup2, "M", item.P41风口中套损坏数);
                        ws.Cells.set_Item(row + rowgroup2, "N", item.P42风口小套损坏数);
                        ws.Cells.set_Item(row + rowgroup2, "O", item.P43渣口大套损坏数);
                        ws.Cells.set_Item(row + rowgroup2, "P", item.P44渣口中套损坏数);
                        ws.Cells.set_Item(row + rowgroup2, "Q", item.P45渣口小套损坏数);
                        ws.Cells.set_Item(row + rowgroup2, "R", item.P46渣铁比);
                        ws.Cells.set_Item(row + rowgroup2, "S", item.P47灰铁比);
                        ws.Cells.set_Item(row + rowgroup2, "T", item.P48原料总耗);
                        ws.Cells.set_Item(row + rowgroup2, "U", item.P49原料单耗);
                        ws.Cells.set_Item(row + rowgroup2, "V", item.P50机烧消耗);
                        ws.Cells.set_Item(row + rowgroup2, "W", item.P51竖炉球消耗);
                        ws.Cells.set_Item(row + rowgroup2, "X", item.P53其它熟料消耗);
                        ws.Cells.set_Item(row + rowgroup2, "Y", item.P54本溪矿消耗);
                        ws.Cells.set_Item(row + rowgroup2, "Z", item.P55其它生料消耗);
                        ws.Cells.set_Item(row + rowgroup2, "AA", item.P56废铁总耗);
                        ws.Cells.set_Item(row + rowgroup2, "AB", item.P57废铁单耗);
                        ws.Cells.set_Item(row + rowgroup2, "AC", item.P58金属收得率);

                        ws.Cells.set_Item(row + rowgroup3, "D", item.P59综合焦炭总耗);
                        ws.Cells.set_Item(row + rowgroup3, "E", item.综合焦炭单耗);
                        ws.Cells.set_Item(row + rowgroup3, "F", item.P82工艺称焦比);
                        ws.Cells.set_Item(row + rowgroup3, "G", item.P60七号称);
                        ws.Cells.set_Item(row + rowgroup3, "H", item.P61干毛焦);
                        ws.Cells.set_Item(row + rowgroup3, "I", item.P62干焦粉);
                        ws.Cells.set_Item(row + rowgroup3, "J", item.P63入炉焦炭总耗);
                        ws.Cells.set_Item(row + rowgroup3, "K", item.P64入炉焦炭单耗);
                        ws.Cells.set_Item(row + rowgroup3, "L", item.P65煤粉总耗);
                        ws.Cells.set_Item(row + rowgroup3, "M", item.煤粉单耗);
                        ws.Cells.set_Item(row + rowgroup3, "N", item.P66焦丁总耗);
                        ws.Cells.set_Item(row + rowgroup3, "O", item.P67焦丁单耗);
                        ws.Cells.set_Item(row + rowgroup3, "P", item.P68燃料比);
                        ws.Cells.set_Item(row + rowgroup3, "Q", item.P69铁成分SI);
                        ws.Cells.set_Item(row + rowgroup3, "R", item.P70铁成分MN);
                        ws.Cells.set_Item(row + rowgroup3, "S", item.P71铁成分P);
                        ws.Cells.set_Item(row + rowgroup3, "T", item.P72铁成分S);
                        ws.Cells.set_Item(row + rowgroup3, "U", item.P73渣成分CAO);
                        ws.Cells.set_Item(row + rowgroup3, "V", item.P74渣成分SIO2);
                        ws.Cells.set_Item(row + rowgroup3, "W", item.P75渣成分AL2O3);
                        ws.Cells.set_Item(row + rowgroup3, "X", item.P76渣成分MGO);
                        ws.Cells.set_Item(row + rowgroup3, "Y", item.P78渣成分S);
                        ws.Cells.set_Item(row + rowgroup3, "Z", item.P79含SI偏差CAO);
                        ws.Cells.set_Item(row + rowgroup3, "AA", item.P80含SI偏差制铁量);
                        ws.Cells.set_Item(row + rowgroup3, "AB", item.P81含SI偏差铸造铁);
                        row++;
                    }

                    wb.Save();
                    wb.Saved = true;
                    app.Visible = true;
                    this.Cursor = Cursors.Default;
                }
            }
        }

        private void SaveToExcel()
        {
             JiShuYueBaoTable table = this.技术月报内容BindingSource.DataSource as JiShuYueBaoTable;
             if (table != null)
             {
                 int nian = Convert.ToInt32(this.TextBox_nian.Text);
                 int yue = Convert.ToInt32(this.TextBox_yue.Text);
                 string title = string.Format("炼铁技术月报{0}年{1}月", nian, yue);
                 string fn = title + ".xls";

                 saveFileDialog1.FileName = fn;
                 if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                 {
                     this.Cursor = Cursors.WaitCursor;
                     
                     string excelfile = saveFileDialog1.FileName;

                     MemoryStream ms = new MemoryStream(Properties.Resources.yuebao_template);

                     HSSFWorkbook hssfworkbook = new HSSFWorkbook(ms);

                     //create a entry of DocumentSummaryInformation
                     DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
                     dsi.Company = "zhc";
                     hssfworkbook.DocumentSummaryInformation = dsi;


                     //create a entry of SummaryInformation
                     SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
                     si.Subject = title;
                     hssfworkbook.SummaryInformation = si;

                     ISheet sheet1 = hssfworkbook.GetSheet("Sheet1");
                     sheet1.GetRow(2).GetCell(6).SetCellValue(nian + "年" + yue + "月"); ;

                     int rowgroup2 = 12;
                     int rowgroup3 = 24;
                     int row = 5;
                     foreach (var item in table)
                     {
                         
                         SetSheetValue(sheet1, row, 3, item.P03全铁产量.Value);
                         SetSheetValue(sheet1, row, 4, item.P05制钢铁.Value);
                         SetSheetValue(sheet1, row, 5, item.P06铸造铁.Value);
                         SetSheetValue(sheet1, row, 6, item.P07折合产量.Value);
                         SetSheetValue(sheet1, row, 7, item.P08合格率.Value);
                         SetSheetValue(sheet1, row, 8, item.P09一级品率.Value);
                         SetSheetValue(sheet1, row, 9, item.P10送炼钢优质铁水率);
                         SetSheetValue(sheet1, row, 10, item.P11高炉利用系数);
                         SetSheetValue(sheet1, row, 11, item.P12入炉焦炭冶炼强度);
                         SetSheetValue(sheet1, row, 12, item.P13综合焦炭冶炼强度);
                         SetSheetValue(sheet1, row, 13, item.P14折算综合焦比);
                         SetSheetValue(sheet1, row, 14, item.P15折合入炉焦比);
                         SetSheetValue(sheet1, row, 15, item.P16入炉矿品位);
                         SetSheetValue(sheet1, row, 16, item.P17熟料比);
                         SetSheetValue(sheet1, row, 17, item.P18入炉焦炭负荷);
                         SetSheetValue(sheet1, row, 18, item.P19综合焦炭负荷);
                         SetSheetValue(sheet1, row, 19, item.P20休风时间);
                         SetSheetValue(sheet1, row, 20, item.P21休风率);
                         SetSheetValue(sheet1, row, 21, item.P22慢风时间);
                         SetSheetValue(sheet1, row, 22, item.P23慢风率);
                         SetSheetValue(sheet1, row, 23, item.P24煤气成分CO2);
                         SetSheetValue(sheet1, row, 24, item.P25煤气成分计算1);
                         SetSheetValue(sheet1, row, 25, item.P27深料线作业率);
                         SetSheetValue(sheet1, row, 26, item.P28回收率);
                         SetSheetValue(sheet1, row, 27, item.P29冷风流量);
                         SetSheetValue(sheet1, row, 28, item.P30平均风温);
                         SetSheetValue(sheet1, row, 29, item.P32热风压力);
                         SetSheetValue(sheet1, row, 30, item.P33炉顶压力);

                         SetSheetValue(sheet1, row + rowgroup2, 3, item.P34炉顶温度);
                         SetSheetValue(sheet1, row + rowgroup2, 4, item.P35富氧率);
                         SetSheetValue(sheet1, row + rowgroup2, 5, item.P36高压率);
                         SetSheetValue(sheet1, row + rowgroup2, 6, item.P37悬料次数);
                         SetSheetValue(sheet1, row + rowgroup2, 7, item.P38坐料次数);
                         SetSheetValue(sheet1, row + rowgroup2, 8, item.P39崩料次数);
                         SetSheetValue(sheet1, row + rowgroup2, 9, item.P40风口大套损坏数);
                         SetSheetValue(sheet1, row + rowgroup2, 10, item.P41风口中套损坏数);
                         SetSheetValue(sheet1, row + rowgroup2, 11, item.P42风口小套损坏数);
                         SetSheetValue(sheet1, row + rowgroup2, 12, item.P43渣口大套损坏数);
                         SetSheetValue(sheet1, row + rowgroup2, 13, item.P44渣口中套损坏数);
                         SetSheetValue(sheet1, row + rowgroup2, 14, item.P45渣口小套损坏数);
                         SetSheetValue(sheet1, row + rowgroup2, 15, item.P46渣铁比);
                         SetSheetValue(sheet1, row + rowgroup2, 16, item.P47灰铁比);
                         SetSheetValue(sheet1, row + rowgroup2, 17, item.P48原料总耗);
                         SetSheetValue(sheet1, row + rowgroup2, 18, item.P49原料单耗);
                         SetSheetValue(sheet1, row + rowgroup2, 19, item.P50机烧消耗);
                         SetSheetValue(sheet1, row + rowgroup2, 20, item.P51竖炉球消耗);
                         SetSheetValue(sheet1, row + rowgroup2, 21, item.PB块);
                         SetSheetValue(sheet1, row + rowgroup2, 22, item.纽曼块);
                         SetSheetValue(sheet1, row + rowgroup2, 23, item.锰矿);
                         SetSheetValue(sheet1, row + rowgroup2, 24, item.钛球);
                         SetSheetValue(sheet1, row + rowgroup2, 25, item.溶剂总耗);
                         SetSheetValue(sheet1, row + rowgroup2, 26, item.溶剂单耗);
                         SetSheetValue(sheet1, row + rowgroup2, 27, item.硅石);
                         SetSheetValue(sheet1, row + rowgroup2, 28, item.白云石);
                         SetSheetValue(sheet1, row + rowgroup2, 29, item.蛇纹石);
                         SetSheetValue(sheet1, row + rowgroup2, 30, item.P58金属收得率);

                         SetSheetValue(sheet1, row + rowgroup3, 3, item.P59综合焦炭总耗);
                         SetSheetValue(sheet1, row + rowgroup3, 4, item.综合焦炭单耗);
                         SetSheetValue(sheet1, row + rowgroup3, 5, item.P82工艺称焦比);
                         SetSheetValue(sheet1, row + rowgroup3, 6, item.P60七号称);
                         SetSheetValue(sheet1, row + rowgroup3, 7, item.P61干毛焦);
                         SetSheetValue(sheet1, row + rowgroup3, 8, item.P62干焦粉);
                         SetSheetValue(sheet1, row + rowgroup3, 9, item.P63入炉焦炭总耗);
                         SetSheetValue(sheet1, row + rowgroup3, 10, item.P64入炉焦炭单耗);
                         SetSheetValue(sheet1, row + rowgroup3, 11, item.P65煤粉总耗);
                         SetSheetValue(sheet1, row + rowgroup3, 12, item.煤粉单耗);
                         SetSheetValue(sheet1, row + rowgroup3, 13, item.P66焦丁总耗);
                         SetSheetValue(sheet1, row + rowgroup3, 14, item.P67焦丁单耗);
                         SetSheetValue(sheet1, row + rowgroup3, 15, item.P68燃料比);
                         SetSheetValue(sheet1, row + rowgroup3, 16, item.P69铁成分SI);
                         SetSheetValue(sheet1, row + rowgroup3, 17, item.P70铁成分MN);
                         SetSheetValue(sheet1, row + rowgroup3, 18, item.P71铁成分P);
                         SetSheetValue(sheet1, row + rowgroup3, 19, item.P72铁成分S);
                         SetSheetValue(sheet1, row + rowgroup3, 20, item.P73渣成分CAO);
                         SetSheetValue(sheet1, row + rowgroup3, 21, item.P74渣成分SIO2);
                         SetSheetValue(sheet1, row + rowgroup3, 22, item.P75渣成分AL2O3);
                         SetSheetValue(sheet1, row + rowgroup3, 23, item.P76渣成分MGO);
                         SetSheetValue(sheet1, row + rowgroup3, 24, item.P78渣成分S);
                         SetSheetValue(sheet1, row + rowgroup3, 25, item.P79含SI偏差CAO);
                         SetSheetValue(sheet1, row + rowgroup3, 26, item.P80含SI偏差制铁量);
                         SetSheetValue(sheet1, row + rowgroup3, 27, item.P81含SI偏差铸造铁);
                         SetSheetValue(sheet1, row + rowgroup3, 28, item.P56废铁总耗);
                         SetSheetValue(sheet1, row + rowgroup3, 29, item.P57废铁单耗);
                         row++;
                     }

                     //Force excel to recalculate all the formula while open
                     //sheet1.ForceFormulaRecalculation = true;

                     FileStream xlsfile = new FileStream(excelfile, FileMode.Create);
                     hssfworkbook.Write(xlsfile);
                     xlsfile.Close();
                     ms.Close();

                     this.Cursor = Cursors.Default;
                 }
             }
        }

        private void SetSheetValue(ISheet sheet, int row, int col, double? val)
        {
            if (val.HasValue) sheet.GetRow(row).GetCell(col).SetCellValue(val.Value);
        }
    }
}
