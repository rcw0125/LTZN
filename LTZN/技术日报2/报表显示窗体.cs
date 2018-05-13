using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Reporting.WinForms;
using System.Data.OracleClient;
using System.IO;
using System.Runtime.Serialization.Formatters.Binary;
using System.Runtime.Serialization;
using Microsoft.Office.Interop.Excel;

namespace LTZN.技术日报
{
    public partial class 报表显示窗体 : Form
    {
        private 技术日报内容 dataSource = null;
        private string[] xiufeng = new string[9];
        private string[] manfeng = new string[9];
        private string[] qitaqingkuang = new string[2];

        public 报表显示窗体()
        {
            InitializeComponent();
        }

        private void 报表显示窗体_Load(object sender, EventArgs e)
        {
                 
        }

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
            madp.FillByRq(xmf.慢风,n.riqi);
          //  qadp.FillByRq(xmf.全厂日消耗, n.riqi);


            foreach (技术日报内容项 x in n)
            {
                x.Convert0ToNull();
            }
            this.技术日报BindingSource.DataSource = n;
            dataSource = n;

            List<ReportParameter> para = new List<ReportParameter>();

            para.Add(new ReportParameter("Riqi",n.riqi.ToLongDateString()));
            int i = 0;
            foreach (XiaohaoGaolu.DataSetXh1.休风Row r in xmf.休风)
            {
                i++;
                string str=i+"、"+r.高炉 +"炉"+r.时间.ToString("H:mm")+"分休风　"+r.间隔.ToString()+"分钟,"+(r.Is原因Null()?"":r.原因);
                if(i<=9)
                {
                  para.Add(new ReportParameter("xiufeng"+i,str));
                  xiufeng[i-1] = str;
                }
            }
            i = 0;
            foreach (XiaohaoGaolu.DataSetXh1.慢风Row r in xmf.慢风)
            {
                i++;
                string str = i + "、" + r.高炉 + "炉" + r.时间.ToString("H:mm") + "分慢风　" + r.间隔.ToString() + "小时," + (r.Is原因Null() ? "" : r.原因);
                if (i <= 9)
                {
                    para.Add(new ReportParameter("manfeng" + i, str));
                    manfeng[i-1] = str;
                }
            }

            OracleConnection cn = new OracleConnection();
            cn.ConnectionString = Properties.Settings.Default.ConnectionString;
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
                    string qtStr=dr.GetString(0);
                    para.Add(new ReportParameter("qita1",qtStr));
                    qitaqingkuang[0] = qtStr;
                }
                if (!dr.IsDBNull(1) && dr.GetString(1) != "")
                {
                    string qtStr = dr.GetString(1);
                    para.Add(new ReportParameter("qita2", qtStr));
                    qitaqingkuang[1] = qtStr;
                }

            }
            dr.Close();
            cn.Close();

            try
            {
                this.reportViewer1.LocalReport.SetParameters(para);
            }
            catch(Exception ex)
            {
                Console.Write(ex);
            }

            this.reportViewer1.RefreshReport();
            
        }

        private void mnuToExcel_Click(object sender, EventArgs e)
        {
            if (dataSource == null)
                throw new Exception("没有数据！");
            this.Cursor = Cursors.WaitCursor;

            ApplicationClass app = new ApplicationClass();
            app.Visible = false;
            Workbook wb = app.Workbooks.Add(Environment.CurrentDirectory + @"\炼铁技术日报.xls");
            wb.Title = "炼铁技术日报" + dataSource.riqi.ToString("yyyy年MM月dd日");
            string fn = wb.Title + ".xls";
            Worksheet ws = (Worksheet)wb.Sheets[1];

            ws.Cells.set_Item(3, "AG", dataSource.riqi.ToString("yyyy年MM月dd日"));

            int row=6;
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
                ws.Cells.set_Item(row, "BF", x.P16原料矿本溪矿);
                ws.Cells.set_Item(row, "BI", x.P17原料矿其它生料);
                ws.Cells.set_Item(row, "BL", x.P18废铁总耗);
                ws.Cells.set_Item(row, "BP", x.P19废铁单耗);
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

                ws.Cells.set_Item(row+15, "B", x.P01单位);
                ws.Cells.set_Item(row + 15, "F", x.P02项目);
                ws.Cells.set_Item(row + 15, "J", x.P31焦丁总耗);
                ws.Cells.set_Item(row + 15, "L", x.P32焦丁单耗);
                ws.Cells.set_Item(row + 15, "O", x.P33综合焦炭总耗);
                ws.Cells.set_Item(row + 15, "R", x.P34综合焦炭单耗);
                ws.Cells.set_Item(row + 15, "V", x.P35综合折算焦比);
                ws.Cells.set_Item(row + 15, "AA", x.P36冶炼强度);
                ws.Cells.set_Item(row + 15, "AE", x.P37焦炭负荷);
                ws.Cells.set_Item(row + 15, "AI", x.P38干毛焦);
                ws.Cells.set_Item(row + 15, "AO", x.P39炼钢铁SI);
                ws.Cells.set_Item(row + 15, "AR", x.P40炼钢铁MN);
                ws.Cells.set_Item(row + 15, "AS", x.P41炼钢铁P);
                ws.Cells.set_Item(row + 15, "AV", x.P42炼钢铁S);
                ws.Cells.set_Item(row + 15, "AZ", x.P43铸造铁SI);
                ws.Cells.set_Item(row + 15, "BB", x.P44铸造铁MN);
                ws.Cells.set_Item(row + 15, "BC", x.P45铸造铁P);
                ws.Cells.set_Item(row + 15, "BE", x.P46铸造铁S);
                ws.Cells.set_Item(row + 15, "BH", x.P47炉渣碱度);
                ws.Cells.set_Item(row + 15, "BK", x.P48休风情况);
                ws.Cells.set_Item(row + 15, "BO", x.P49慢风);
                ws.Cells.set_Item(row + 15, "BR", x.P50坐料次数);
                ws.Cells.set_Item(row + 15, "BW", x.P51悬料次数);
                ws.Cells.set_Item(row + 15, "BZ", x.P52崩料次数);
                ws.Cells.set_Item(row + 15, "CD", x.P53风口损坏数大);
                ws.Cells.set_Item(row + 15, "CG", x.P54风口损坏数中);
                ws.Cells.set_Item(row + 15, "CH", x.P55风口损坏数小);
                ws.Cells.set_Item(row + 15, "CN", x.P56渣口损坏数大);
                ws.Cells.set_Item(row + 15, "CO", x.P57渣口损坏数中);
                ws.Cells.set_Item(row + 15, "CR", x.P58渣口损坏数小);
                ws.Cells.set_Item(row + 15, "CV", x.P59本厂铸造SI大);
                ws.Cells.set_Item(row + 15, "CY", x.P60本厂铸造SI小);
                ws.Cells.set_Item(row + 15, "DF", x.P61送炼钢厂SI大);
                ws.Cells.set_Item(row + 15, "DK", x.P63送炼钢厂SI小);

                row++;
            }

            ws.Cells.set_Item(34, "M", xiufeng[0]);
            ws.Cells.set_Item(35, "M", xiufeng[1]);
            ws.Cells.set_Item(36, "M", xiufeng[2]);
            ws.Cells.set_Item(34, "AZ", xiufeng[3]);
            ws.Cells.set_Item(35, "AZ", xiufeng[4]);
            ws.Cells.set_Item(36, "AZ", xiufeng[5]);
            ws.Cells.set_Item(34, "CJ", xiufeng[6]);
            ws.Cells.set_Item(35, "CJ", xiufeng[7]);
            ws.Cells.set_Item(36, "CJ", xiufeng[8]);

            ws.Cells.set_Item(37, "M", manfeng[0]);
            ws.Cells.set_Item(38, "M", manfeng[1]);
            ws.Cells.set_Item(39, "M", manfeng[2]);
            ws.Cells.set_Item(37, "AZ", manfeng[3]);
            ws.Cells.set_Item(38, "AZ", manfeng[4]);
            ws.Cells.set_Item(39, "AZ", manfeng[5]);
            ws.Cells.set_Item(37, "CJ", manfeng[6]);
            ws.Cells.set_Item(38, "CJ", manfeng[7]);
            ws.Cells.set_Item(39, "CJ", manfeng[8]);

            ws.Cells.set_Item(40, "M", qitaqingkuang[0]);
            ws.Cells.set_Item(41, "M", qitaqingkuang[1]);

            string fullfn = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + fn;
            wb.SaveCopyAs(fullfn);
            wb.Saved = true;
            app.Visible = true;
            this.Cursor = Cursors.Default;
        }

    }
}