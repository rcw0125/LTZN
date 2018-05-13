using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using C1.Win.C1Chart;

namespace LTZN.数据分析
{
    public partial class 物理热趋势 : Form
    {
        public 物理热趋势()
        {
            InitializeComponent();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (r1.RadioChecked && (sjEnd.Value < sjBegin.Value))
            {
                MessageBox.Show("结束时间应该比开始时间大！");
                return;
            }

            c1Chart1.ChartGroups[0].ChartData.SeriesList.Clear();
            int gaolu = 0;
            switch (this.toolStrip_gaolu.Text)
            {
                case "1高炉":
                    gaolu = 1;
                    break;
                case "2高炉":
                    gaolu = 2;
                    break;
                case "3高炉":
                    gaolu = 3;
                    break;
                case "4高炉":
                    gaolu = 4;
                    break;
                case "5高炉":
                    gaolu = 5;
                    break;
                case "6高炉":
                    gaolu = 6;
                    break;
            }

            // Add the data
            ChartData data = c1Chart1.ChartGroups[0].ChartData;
            ChartDataSeriesCollection series = data.SeriesList;
       
            double[] dy;
            double max;
            double min;
            LegendEnviroment.loadWlrSetting(gaolu, out dy, out max, out min);
            c1Chart1.ChartArea.AxisY.Max = max;
            c1Chart1.ChartArea.AxisY.Min = min;
            
            Color[] dyColor = new Color[] { Color.FromArgb(255, 192, 192), Color.FromArgb(255, 255, 192), Color.FromArgb(192, 255, 192), Color.FromArgb(255, 255, 192), Color.FromArgb(255, 192, 192) };
            string[] dyName = new string[] { "过大", "有点大", "正常", "有点小", "过小" };
            
            
            //plot the student scores
            ChartDataSeries StuSeries = series.AddNewSeries();
            StuSeries.Label = "物理热趋势";
            StuSeries.LineStyle.Pattern = LinePatternEnum.Solid;
            StuSeries.LineStyle.Color = Color.DarkBlue;
            StuSeries.SymbolStyle.Shape = SymbolShapeEnum.Star;
            StuSeries.SymbolStyle.Color = Color.DarkRed;
             
            if (r1.RadioChecked)
            {
                c1Chart1.ChartArea.AxisX.AnnoFormat = FormatEnum.DateManual;
                c1Chart1.ChartArea.AxisX.AnnoFormatString = "MM/dd HH:mm";
                c1Chart1.ChartArea.AxisX.AnnotationRotation = 0;
                DateTime[] siT;
                double[] si;
                string wulire = "FEWENDU";
                LegendEnviroment.loadTiO2(gaolu,wulire, sjBegin.Value, sjEnd.Value, out siT, out si);
                label1.Text = "标准方差：" + LegendEnviroment.stddev(si).ToString("###0.000");
                StuSeries.X.CopyDataIn(siT);
                StuSeries.Y.CopyDataIn(si);
                StuSeries = null;

                c1Chart1.Footer.Text = this.toolStrip_gaolu.Text + "  " + this.sjBegin.Text + "  ----  " + this.sjEnd.Text;

                DateTime[] ax = new DateTime[] { sjBegin.Value, sjEnd.Value };

                for (int i = 0; i < 4; i++)
                {
                    double[] ay1 = new double[] { dy[i], dy[i] };
                    StuSeries = series.AddNewSeries();
                    StuSeries.Label = "最大值";
                    StuSeries.LineStyle.Pattern = LinePatternEnum.Solid;
                    StuSeries.LineStyle.Color = Color.Black;
                    StuSeries.SymbolStyle.Shape = SymbolShapeEnum.None;
                    StuSeries.SymbolStyle.Color = Color.DarkRed;
                    StuSeries.X.CopyDataIn(ax);
                    StuSeries.Y.CopyDataIn(ay1);
                    StuSeries = null;
                }

            }
            else
            {
                c1Chart1.ChartArea.AxisX.AnnoFormat = FormatEnum.NumericGeneral;
                c1Chart1.ChartArea.AxisX.AnnotationRotation = 90;
                long[] siLuci;
                double[] si;
                string wulire = "FEWENDU";
                LegendEnviroment.loadTiO2(gaolu,wulire,toolStrip_Luci1.Text, toolStrip_Luci2.Text, out siLuci, out si);
                label1.Text = "标准方差：" + LegendEnviroment.stddev(si).ToString("###0.000");
                StuSeries.X.CopyDataIn(siLuci);
                StuSeries.Y.CopyDataIn(si);
                StuSeries = null;

                c1Chart1.Footer.Text = this.toolStrip_gaolu.Text + "  " + this.toolStrip_Luci1.Text + "  ----  " + this.toolStrip_Luci2.Text;

                long[] ax = new long[] { Convert.ToInt64(toolStrip_Luci1.Text), Convert.ToInt64(toolStrip_Luci2.Text) };

                for (int i = 0; i < 4; i++)
                {
                    double[] ay1 = new double[] { dy[i], dy[i] };
                    StuSeries = series.AddNewSeries();
                    StuSeries.Label = "最大值";
                    StuSeries.LineStyle.Pattern = LinePatternEnum.Solid;
                    StuSeries.LineStyle.Color = Color.Black;
                    StuSeries.SymbolStyle.Shape = SymbolShapeEnum.None;
                    StuSeries.SymbolStyle.Color = Color.DarkRed;
                    StuSeries.X.CopyDataIn(ax);
                    StuSeries.Y.CopyDataIn(ay1);
                    StuSeries = null;
                }
            }

            // Add and show the alarm zones
            AlarmZonesCollection zones = c1Chart1.ChartArea.PlotArea.AlarmZones;
            zones.Clear();
            for (int i = 0; i < 5; i++)
            {
                AlarmZone zone = zones.AddNewZone();

                zone.Name = dyName[i];
                zone.BackColor = dyColor[i];

                if (i == 0)
                    zone.UpperExtent = c1Chart1.ChartArea.AxisY.Max;
                else
                    zone.UpperExtent = zones[i - 1].LowerExtent;

                if (i == 4)
                    zone.LowerExtent = c1Chart1.ChartArea.AxisY.Min;
                else
                    zone.LowerExtent = dy[i];

                zone.Visible = true;
            }

        }

        private void 炉温_Load(object sender, EventArgs e)
        {
            this.sjBegin.Value = DateTime.Today.AddDays(-2);
            this.sjEnd.Value = DateTime.Now;
        }

        private void r1_CheckedChanged(object sender, EventArgs e)
        {
            if (r1.RadioChecked)
            {
                sjBegin.Visible = true;
                sjEnd.Visible = true;
                toolStrip_Luci1.Visible = false;
                toolStrip_Luci2.Visible = false;
                toolStripLabel1.Text = "开始时间";
                toolStripLabel2.Text = "结束时间";
            }
            else
            {
                sjBegin.Visible = false;
                sjEnd.Visible = false;
                toolStrip_Luci1.Visible = true;
                toolStrip_Luci2.Visible = true;
                toolStripLabel1.Text = "开始炉次";
                toolStripLabel2.Text = "结束炉次";
            }
        }

        private void toolStrip_Luci1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                toolStrip_Luci2.Focus();
        }

        private void toolStrip_Luci2_KeyDown(object sender, KeyEventArgs e)
        {
        }
    }
}