using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using LTZN.Repository;
using System.Data.OracleClient;
using Microsoft.Reporting.WinForms;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace LTZN.质量日报
{
    public partial class 质量日报Form : Form
    {
        public 质量日报Form()
        {
            InitializeComponent();
        }
        private 质量日报 _zhiliangRiBao = null;
        public 质量日报 ZhiliangRiBao
        {
            get
            {
                if (_zhiliangRiBao == null)
                {
                    _zhiliangRiBao = new 质量日报();
                }
                return _zhiliangRiBao;
            }
        }
        private void 质量日报Form_Load(object sender, EventArgs e)
        {
            // TODO: 这行代码将数据加载到表“dDTTY._DDTTY”中。您可以根据需要移动或移除它。
            this.dDTTYTableAdapter.Fill(this.dDTTY._DDTTY);
            //  this.reportViewer1.ProcessingMode = Microsoft.Reporting.WinForms.ProcessingMode.Local;
            //   this.reportViewer1.LocalReport.ReportPath = @"质量日报\质量日报.rdlc";
            //    reportViewer1.LocalReport.DataSources.Add(new ReportDataSource("zlrb",loaddate()));
            riqiControl.Value = DateTime.Today;
            //     this.reportViewer1.RefreshReport();
        }

        private void btnShowReport_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            this.reportViewer1.Visible = true;

            ZhiliangRiReportRepository rp = new ZhiliangRiReportRepository();
            质量日报内容 n = rp.GetRiReport(riqiControl.Value.Date);
            机烧粒度内容 m = rp.GetJSReport(riqiControl.Value.Date);
            燃料指标内容 r = rp.GetRLReport(riqiControl.Value.Date);
            套筒窑内容 t = rp.GetTTYReport(riqiControl.Value.Date);

            setDataSource(n, m, r,t);
            this.Cursor = Cursors.Default;
        }
        质量日报内容 质量日报内容dataSource = null;
        机烧粒度内容 机烧粒度内容dataSource = null;
        燃料指标内容 燃料指标内容dataSource = null;
        套筒窑内容 套筒窑内容dataSource = null;
        public void setDataSource(质量日报内容 n, 机烧粒度内容 m, 燃料指标内容 r,套筒窑内容 t)
        {
            质量日报内容dataSource = n;
            机烧粒度内容dataSource = m;
            燃料指标内容dataSource = r;
            this.dDJSBindingSource.DataSource = m;
            this.DDYLBindingSource.DataSource = n;
            this.燃料指标内容项BindingSource.DataSource = r;
            List<ReportParameter> para = new List<ReportParameter>();

            para.Add(new ReportParameter("Riqi", n.riqi.ToString("yyyy年MM月dd日")));
            String leader = LTZN.Data.SysParameter.GetStrParameter("单位负责人", "郝江敏");
            String dayReportAuthor = LTZN.Data.SysParameter.GetStrParameter("日报负责人", "刘红芬");
            String monthReportAuthor = LTZN.Data.SysParameter.GetStrParameter("月报负责人", "王文英");


            string yejiao = "单位负责人：“" + leader + "”　 制表人：“" + dayReportAuthor + "”　 第 ＋　＋ 页，共一页";
            para.Add(new ReportParameter("yejiao", yejiao));
            para.Add(new ReportParameter("actUser", dayReportAuthor));
            this.reportViewer1.LocalReport.SetParameters(para);
            this.reportViewer1.RefreshReport();
            this.reportViewer1.Refresh();

        }



        private void btnSave_Click(object sender, EventArgs e)
        {
         //   this.Cursor = Cursors.WaitCursor;
        //    DateTime riqi = riqiControl.Value.Date;
       //     List<string> items = new List<string>();
        //    OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
        //    conn.Open();
         //   OracleCommand cmd = new OracleCommand("Select MC,SJ from DDYL  Where trunc(SJ)= :RQ)", conn);
         //   cmd.Parameters.Add(":RQ", OracleType.DateTime).Value = riqi;
        //    OracleDataReader dr = cmd.ExecuteReader();
        //    while (dr.Read())
          //  {
         //       if (!dr.IsDBNull(0) && !dr.IsDBNull(1))
          //      {
        //            items.Add(dr.GetString(0) + "-" + dr.GetString(1));
         //       }
          //  }
          //  dr.Close();
         //   conn.Close();
        }

        private void btnQuery_Click(object sender, EventArgs e)
        {

            this.Cursor = Cursors.WaitCursor;
            this.reportViewer1.Visible = true;

            ZhiliangRiReportRepository rp = new ZhiliangRiReportRepository();
            质量日报内容 n = rp.GetRiReport(riqiControl.Value.Date);
            机烧粒度内容 m = rp.GetJSReport(riqiControl.Value.Date);
            燃料指标内容 r = rp.GetRLReport(riqiControl.Value.Date);
            套筒窑内容 t = rp.GetTTYReport(riqiControl.Value.Date);
            setDataSource(n, m, r,t);
            this.Cursor = Cursors.Default;
        }

        private void btnModel_Click(object sender, EventArgs e)
        {
            if (this.ZhiliangRiBao.myModel != null)
            {
                CalModelForm form = new CalModelForm(this.ZhiliangRiBao.myModel);
                form.ShowDialog();
                this.ZhiliangRiBao.display(webBrowser1.Document);
            }
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            if (reportViewer1.Visible)
                reportViewer1.PrintDialog();

        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            string title = "炼铁质量日报" + 质量日报内容dataSource.riqi.ToString("yyyy年MM月dd日");
            string fn = title + ".xls";
            saveFileDialog1.FileName = fn;
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {

                this.Cursor = Cursors.WaitCursor;
                Stream file = saveFileDialog1.OpenFile();
                file.Write(Properties.Resources.炼铁质量日报, 0, Properties.Resources.炼铁质量日报.Length);
                file.Close();

                object m_objOpt = System.Reflection.Missing.Value;

                ApplicationClass app = new ApplicationClass();
                app.Visible = false;
                Workbook wb = app.Workbooks.Open(saveFileDialog1.FileName, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt, m_objOpt);
                wb.Title = title;
                Worksheet ws = (Worksheet)wb.Sheets[1];

                ws.Cells.set_Item(3, "E", 质量日报内容dataSource.riqi.ToString("yyyy年MM月dd日"));

                int row = 7;
                if (质量日报内容dataSource != null)
                {
                    foreach (质量日报内容项 x in 质量日报内容dataSource)
                    {
                        ws.Cells.set_Item(row, "A", x.MC);
                        ws.Cells.set_Item(row, "B", x.TFE);
                        ws.Cells.set_Item(row, "C", x.FeO);
                        ws.Cells.set_Item(row, "D", x.SiO2);
                        ws.Cells.set_Item(row, "E", x.CaO);
                        ws.Cells.set_Item(row, "F", x.MgO);
                        ws.Cells.set_Item(row, "G", x.Al2O3);
                        ws.Cells.set_Item(row, "H", x.S);
                        ws.Cells.set_Item(row, "I", x.P);
                        ws.Cells.set_Item(row, "J", x.TiO2);
                        ws.Cells.set_Item(row, "K", x.R);
                        ws.Cells.set_Item(row, "L", x.Cr);
                        ws.Cells.set_Item(row, "M", x.Ni);
                        ws.Cells.set_Item(row, "N", x.As);
                        ws.Cells.set_Item(row, "O", x.MnO);
                        ws.Cells.set_Item(row, "P", x.Pb);
                        ws.Cells.set_Item(row, "Q", x.Zn);
                        ws.Cells.set_Item(row, "R", x.JJS);
                        ws.Cells.set_Item(row, "S", x.V2O5);
                        ws.Cells.set_Item(row, "T", x.Cu);
                        ws.Cells.set_Item(row, "U", x.Mo);
                        ws.Cells.set_Item(row, "V", x.Sn);
                        ws.Cells.set_Item(row, "W", x.Sb);

                        row++;
                    }

                }


                int row1 = 19;
                if (机烧粒度内容dataSource != null)
                {

                    foreach (机烧粒度内容项 x in 机烧粒度内容dataSource)
                    {
                        ws.Cells.set_Item(row1, "A", x.MC);
                        ws.Cells.set_Item(row1, "B", x.ZG);
                        ws.Cells.set_Item(row1, "C", x.KM);
                        ws.Cells.set_Item(row1, "D", x.SF);
                        ws.Cells.set_Item(row1, "E", x.KY);
                        ws.Cells.set_Item(row1, "F", x.DY40);
                        ws.Cells.set_Item(row1, "G", x.ZJ425);
                        ws.Cells.set_Item(row1, "H", x.ZJ2516);
                        ws.Cells.set_Item(row1, "I", x.ZJ1610);
                        ws.Cells.set_Item(row1, "J", x.ZJ105);
                        ws.Cells.set_Item(row1, "K", x.XY5);


                        row1++;
                    }

                    int row2 = 29;
                    if (燃料指标内容dataSource != null)
                    {

                        foreach (燃料指标内容项 x in 燃料指标内容dataSource)
                        {
                            ws.Cells.set_Item(row2, "A", x.MC);
                            ws.Cells.set_Item(row2, "B", x.C);
                            ws.Cells.set_Item(row2, "C", x.HUIFEN);
                            ws.Cells.set_Item(row2, "D", x.HUIFA);
                            ws.Cells.set_Item(row2, "E", x.S);
                            ws.Cells.set_Item(row2, "F", x.SF);
                            ws.Cells.set_Item(row2, "G", x.M40);
                            ws.Cells.set_Item(row2, "H", x.M10);
                            ws.Cells.set_Item(row2, "J", x.QD);
                            ws.Cells.set_Item(row2, "K", x.DY80);
                            ws.Cells.set_Item(row2, "L", x.ZJ86);
                            ws.Cells.set_Item(row2, "M", x.ZJ64);
                            ws.Cells.set_Item(row2, "N", x.ZJ42);
                            ws.Cells.set_Item(row2, "O", x.XY25);
                            ws.Cells.set_Item(row2, "P", x.PJLD);
                            ws.Cells.set_Item(row2, "Q", x.LDJYX);

                            row2++;
                        }



                    }
                    int row3 = 38;
                    if (套筒窑内容dataSource != null)
                    {




                        foreach (套筒窑内容项 item in 套筒窑内容dataSource)
                        {
                            ws.Cells.set_Item(row3, "A", item.MC);
                            ws.Cells.set_Item(row3, "A", item.CaO);
                            ws.Cells.set_Item(row3, "A", item.SiO2);
                            ws.Cells.set_Item(row3, "A", item.Al2O3);
                            ws.Cells.set_Item(row3, "A", item.MgO);
                            ws.Cells.set_Item(row3, "A", item.S);
                            ws.Cells.set_Item(row3, "A", item.P);
                            ws.Cells.set_Item(row3, "A", item.HXD);
                            ws.Cells.set_Item(row3, "A", item.ZJ);
                            ws.Cells.set_Item(row3, "A", item.DY60);
                            ws.Cells.set_Item(row3, "A", item.ZJ64);
                            ws.Cells.set_Item(row3, "A", item.XY40);
                            row3++;
                        }
                   
                    
                    
                    }


                    String leader = LTZN.Data.SysParameter.GetStrParameter("单位负责人", "郝江敏");
                    String dayReportAuthor = LTZN.Data.SysParameter.GetStrParameter("日报负责人", "刘红芬");
                    String monthReportAuthor = LTZN.Data.SysParameter.GetStrParameter("月报负责人", "王文英");

                    string yejiao = "单位负责人：" + leader + "“　“制表人： " + dayReportAuthor + "”　”第“＋　＋”页，共一页”";
                    ws.Cells.set_Item(44, "B", yejiao);

                    wb.Save();
                    wb.Saved = true;
                    app.Visible = true;
                    this.Cursor = Cursors.Default;
                }
            }
        }
    }
}