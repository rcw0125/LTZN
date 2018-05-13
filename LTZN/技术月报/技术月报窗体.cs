using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Reporting.WinForms;
using System.Xml;
using System.Xml.Serialization;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace LTZN.技术月报
{
    public partial class 技术月报窗体 : Form
    {
        public 技术月报窗体()
        {
            InitializeComponent();
            
        }

        private void 技术月报窗体_Load(object sender, EventArgs e)
        {
            // TODO: 这行代码将数据加载到表“技术月报表.技术月报”中。您可以根据需要移动或移除它。
           // this.技术月报TableAdapter.FillBy年月(this.技术月报表.技术月报);

           // this.reportViewer1.RefreshReport();
         
        }

        public void 显示报表(int nian, int yue)
        {
            List<ReportParameter> para = new List<ReportParameter>();
            技术月报TableAdapter1.Fill(技术月报表.技术月报, Convert.ToDecimal(nian), Convert.ToDecimal(yue));
            para.Add(new ReportParameter("actUser", Properties.Settings.Default.制表人));
            para.Add(new ReportParameter("actLeader", Properties.Settings.Default.负责人));
            try
            {
                this.reportViewer1.LocalReport.SetParameters(para);
            }
            catch (Exception ex)
            {
                Console.Write(ex);
            }
            this.reportViewer1.RefreshReport();
            this.Show();
        }

        private void reportViewer1_Load(object sender, EventArgs e)
        {

        }

        private void 打开ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                技术月报表.Reset();
                技术月报表.ReadXml(openFileDialog1.FileName);
                this.reportViewer1.RefreshReport();
            }
        }

        private void 保存ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                技术月报表.WriteXml(saveFileDialog1.FileName);
            }
        }

        private void 导出到ExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (this.技术月报表.技术月报.Rows.Count != 12)
                throw new Exception("数据源格式不对！");
          
            int nian = Convert.ToInt32(this.技术月报表.技术月报.Rows[0]["P年"].ToString());
            int yue = Convert.ToInt32(this.技术月报表.技术月报.Rows[0]["P月"].ToString());
            string title = "炼铁技术月报" + nian + "年" + yue + "月";
            string fn = title + ".xls";
            saveFileDialog2.FileName = fn;
            if (saveFileDialog2.ShowDialog() == DialogResult.OK)
            {
                this.Cursor = Cursors.WaitCursor;

                Stream file = saveFileDialog2.OpenFile();
                file.Write(Properties.Resources.炼铁技术月报, 0, Properties.Resources.炼铁技术月报.Length);
                file.Close();

                object m_objOpt = System.Reflection.Missing.Value;

                ApplicationClass app = new ApplicationClass();
                app.Visible = false;
                Workbook wb = app.Workbooks.Open(saveFileDialog2.FileName, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt);
                wb.Title = title;
                Worksheet ws = (Worksheet)wb.Sheets[1];

                fillExcel(nian, yue, this.技术月报表.技术月报, ws, Properties.Settings.Default.制表人);

                wb.Save();
                wb.Saved = true;
                app.Visible = true;
                this.Cursor = Cursors.Default;
            }
        }
      
        private void fillExcel(int year, int yue, System.Data.DataTable dt, Worksheet ws, string author)
        {

            if (dt.Rows.Count != 12)
                throw new Exception("数据源格式不对！");
                

            ws.Cells[3, 7] = year + "年" + yue + "月";
            ws.Cells[3, 27] = author;

            int ExcelbeginRow = 6;
            int ExcelbeginCol = 4;

            List<string> cols = new List<string>();
            cols.Add("P03全铁产量");
            cols.Add("P05制钢铁");
            cols.Add("P06铸造铁");
            cols.Add("P07折合产量");
            cols.Add("P08合格率");
            cols.Add("P09一级品率");
            cols.Add("P10送炼钢优质铁水率");
            cols.Add("P11高炉利用系数");
            cols.Add("P12入炉焦炭冶炼强度");
            cols.Add("P13综合焦炭冶炼强度");
            cols.Add("P14折算综合焦比");
            cols.Add("P15折合入炉焦比");
            cols.Add("P16入炉矿品位");
            cols.Add("P17熟料比");
            cols.Add("P18入炉焦炭负荷");
            cols.Add("P19综合焦炭负荷");
            cols.Add("P20休风时间");
            cols.Add("P21休风率");
            cols.Add("P22慢风时间");
            cols.Add("P23慢风率");
            cols.Add("P24煤气成分CO2");
            cols.Add("P25煤气成分计算1");
            cols.Add("P27深料线作业率");
            cols.Add("P28回收率");
            cols.Add("P29冷风流量");
            cols.Add("P30平均风温");
            cols.Add("P32热风压力");
            cols.Add("P33炉顶压力");
            cols.Add("P34炉顶温度");
            cols.Add("P35富氧率");
            cols.Add("P36高压率");
            cols.Add("P37悬料次数");
            cols.Add("P38坐料次数");
            cols.Add("P39崩料次数");
            cols.Add("P40风口大套损坏数");
            cols.Add("P41风口中套损坏数");
            cols.Add("P42风口小套损坏数");
            cols.Add("P43渣口大套损坏数");
            cols.Add("P44渣口中套损坏数");
            cols.Add("P45渣口小套损坏数");
            cols.Add("P46渣铁比");
            cols.Add("P47灰铁比");
            cols.Add("P48原料总耗");
            cols.Add("P49原料单耗");
            cols.Add("P50机烧消耗");
            cols.Add("P51竖炉球消耗");
            cols.Add("P53其它熟料消耗");
            cols.Add("P54本溪矿消耗");
            cols.Add("P55其它生料消耗");
            cols.Add("P56废铁总耗");
            cols.Add("P57废铁单耗");
            cols.Add("P58金属收得率");
            cols.Add("P59综合焦炭总耗");
            cols.Add("综合焦炭单耗");
            cols.Add("P82工艺称焦比");
            cols.Add("P60七号称");
            cols.Add("P61干毛焦");
            cols.Add("P62干焦粉");
            cols.Add("P63入炉焦炭总耗");
            cols.Add("P64入炉焦炭单耗");
            cols.Add("P65煤粉总耗");
            cols.Add("P66焦丁总耗");
            cols.Add("P67焦丁单耗");
            cols.Add("P68燃料比");
            cols.Add("P69铁成分SI");
            cols.Add("P70铁成分MN");
            cols.Add("P71铁成分P");
            cols.Add("P72铁成分S");
            cols.Add("P73渣成分CAO");
            cols.Add("P74渣成分SIO2");
            cols.Add("P75渣成分AL2O3");
            cols.Add("P76渣成分MGO");
            cols.Add("P77渣成分FEO");
            cols.Add("P78渣成分S");
            cols.Add("P79含SI偏差CAO");
            cols.Add("P80含SI偏差制铁量");
            cols.Add("P81含SI偏差铸造铁");
            for (int i = 0; i < 12; i++)
            {
                for (int j = 0; j < cols.Count; j++)
                {
                    if (dt.Columns.Contains(cols[j]))
                        ws.Cells[ExcelbeginRow + Math.Floor((double)(j / 26)) * 14 + i, ExcelbeginCol + j % 26] = dt.Rows[i][cols[j]];
                }
            }
        }
    }
}