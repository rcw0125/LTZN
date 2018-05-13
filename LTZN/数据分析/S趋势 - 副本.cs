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
    public partial class TiO2趋势 : Form
    {
        public TiO2趋势()
        {
            InitializeComponent();
        }

        private void S趋势_Load(object sender, EventArgs e)
        {
            splitContainer1.SplitterDistance = 200;
            this.sjBegin.Value = DateTime.Today.AddDays(-2);
            this.sjEnd.Value = DateTime.Now;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (r1.RadioChecked && (sjEnd.Value < sjBegin.Value))
            {
                MessageBox.Show("结束时间应该比开始时间大！");
                return;
            }
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

            c1Chart1.ChartGroups[0].ChartData.SeriesList.Clear();
            c1Chart2.ChartGroups[0].ChartData.SeriesList.Clear();

            // Add the data
            ChartData data = c1Chart1.ChartGroups[0].ChartData;
            ChartDataSeriesCollection series = data.SeriesList;

            ChartData data2 = c1Chart2.ChartGroups[0].ChartData;
            ChartDataSeriesCollection series2 = data2.SeriesList;

            double[] dy1;
            double[] dy2;
            double max;
            double min;
            double max2;
            double min2;
            LegendEnviroment.loadSSetting(out dy1, out dy2, out max, out min, out max2, out min2);
            c1Chart1.ChartArea.AxisY.Max = max;
            c1Chart1.ChartArea.AxisY.Min = min;

            c1Chart2.ChartArea.AxisY.Max = max2;
            c1Chart2.ChartArea.AxisY.Min = min2;

            Color[] dyColor = new Color[] { Color.FromArgb(255, 192, 192),  Color.FromArgb(255, 255, 192), Color.FromArgb(192, 255, 192) };
            string[] dyName = new string[] { "过大", "正常", "过小" };


            //plot the student scores
            ChartDataSeries StuSeries = series.AddNewSeries();
            StuSeries.Label = "S(炉温)趋势";
            StuSeries.LineStyle.Pattern = LinePatternEnum.Solid;
            StuSeries.LineStyle.Color = Color.DarkBlue;
            StuSeries.SymbolStyle.Shape = SymbolShapeEnum.Star;
            StuSeries.SymbolStyle.Color = Color.DarkRed;

            //plot the student scores
            ChartDataSeries StuSeries2 = series2.AddNewSeries();
            StuSeries2.Label = "S(炉温)趋势";
            StuSeries2.LineStyle.Pattern = LinePatternEnum.Solid;
            StuSeries2.LineStyle.Color = Color.DarkBlue;
            StuSeries2.SymbolStyle.Shape = SymbolShapeEnum.Star;
            StuSeries2.SymbolStyle.Color = Color.DarkRed;

            if (r1.RadioChecked)
            {
                c1Chart1.ChartArea.AxisX.AnnoFormat = FormatEnum.DateManual;
                c1Chart1.ChartArea.AxisX.AnnoFormatString = "MM/dd HH:mm";
                c1Chart1.ChartArea.AxisX.AnnotationRotation = 0;
                c1Chart2.ChartArea.AxisX.AnnoFormat = FormatEnum.DateManual;
                c1Chart2.ChartArea.AxisX.AnnoFormatString = "MM/dd HH:mm";
                c1Chart2.ChartArea.AxisX.AnnotationRotation = 0;
                DateTime[] sT;
                double[] s;
                DateTime[] sTCha;
                double[] sCha;
                LegendEnviroment.loadS(gaolu, sjBegin.Value, sjEnd.Value, out sT, out s);
                StuSeries.X.CopyDataIn(sT);
                StuSeries.Y.CopyDataIn(s);
                StuSeries = null;
               
                sCha = calJicha(s);
                if (sCha != null)
                {
                    sTCha = new DateTime[sT.Length - 1];
                    for (int i = 1; i < sT.Length; i++)
                    {
                        sTCha[i - 1] = sT[i];
                    }
                    StuSeries2.X.CopyDataIn(sT);
                    StuSeries2.Y.CopyDataIn(sCha);
                }
                StuSeries2 = null;

                c1Chart2.Footer.Text = this.toolStrip_gaolu.Text + "  " + this.sjBegin.Text + "  ----  " + this.sjEnd.Text;

                DateTime[] ax = new DateTime[] { sjBegin.Value, sjEnd.Value };

                for (int i = 0; i < 3; i++)
                {
                    double[] ay1 = new double[] { dy1[i], dy1[i] };
                    StuSeries = series.AddNewSeries();
                    StuSeries.Label = i.ToString();
                    StuSeries.LineStyle.Pattern = LinePatternEnum.Solid;
                    StuSeries.LineStyle.Color = Color.Black;
                    StuSeries.SymbolStyle.Shape = SymbolShapeEnum.None;
                    StuSeries.SymbolStyle.Color = Color.DarkRed;
                    StuSeries.X.CopyDataIn(ax);
                    StuSeries.Y.CopyDataIn(ay1);
                    StuSeries = null;

                    double[] ay2 = new double[] { dy2[i], dy2[i] };
                    StuSeries2 = series2.AddNewSeries();
                    StuSeries2.Label = i.ToString();
                    StuSeries2.LineStyle.Pattern = LinePatternEnum.Solid;
                    StuSeries2.LineStyle.Color = Color.Black;
                    StuSeries2.SymbolStyle.Shape = SymbolShapeEnum.None;
                    StuSeries2.SymbolStyle.Color = Color.DarkRed;
                    StuSeries2.X.CopyDataIn(ax);
                    StuSeries2.Y.CopyDataIn(ay2);
                    StuSeries2 = null;
                }

            }
            else
            {
                c1Chart1.ChartArea.AxisX.AnnoFormat = FormatEnum.NumericGeneral;
                c1Chart1.ChartArea.AxisX.AnnotationRotation = 90;
                c1Chart2.ChartArea.AxisX.AnnoFormat = FormatEnum.NumericGeneral;
                c1Chart2.ChartArea.AxisX.AnnotationRotation = 90;
                long[] sLuci;
                double[] s;
                LegendEnviroment.loadS(gaolu, toolStrip_Luci1.Text, toolStrip_Luci2.Text, out sLuci, out s);
                StuSeries.X.CopyDataIn(sLuci);
                StuSeries.Y.CopyDataIn(s);
                StuSeries = null;

                long[] sLuciCha;
                double[] sCha;
                sCha = calJicha(s);
                if (sCha != null)
                {
                    sLuciCha = new long[sLuci.Length - 1];
                    for (int i = 1; i < sLuci.Length; i++)
                    {
                        sLuciCha[i - 1] = sLuci[i];
                    }
                    StuSeries2.X.CopyDataIn(sLuciCha);
                    StuSeries2.Y.CopyDataIn(sCha);
                }
                StuSeries2 = null;

                c1Chart2.Footer.Text = this.toolStrip_gaolu.Text + "  " + this.toolStrip_Luci1.Text + "  ----  " + this.toolStrip_Luci2.Text;

                long[] ax = new long[] { Convert.ToInt64(toolStrip_Luci1.Text), Convert.ToInt64(toolStrip_Luci2.Text) };

                for (int i = 0; i < 3; i++)
                {
                    double[] ay1 = new double[] { dy1[i], dy1[i] };
                    StuSeries = series.AddNewSeries();
                    StuSeries.Label = i.ToString();
                    StuSeries.LineStyle.Pattern = LinePatternEnum.Solid;
                    StuSeries.LineStyle.Color = Color.Black;
                    StuSeries.SymbolStyle.Shape = SymbolShapeEnum.None;
                    StuSeries.SymbolStyle.Color = Color.DarkRed;
                    StuSeries.X.CopyDataIn(ax);
                    StuSeries.Y.CopyDataIn(ay1);
                    StuSeries = null;

                    double[] ay2 = new double[] { dy2[i], dy2[i] };
                    StuSeries2 = series2.AddNewSeries();
                    StuSeries2.Label = i.ToString();
                    StuSeries2.LineStyle.Pattern = LinePatternEnum.Solid;
                    StuSeries2.LineStyle.Color = Color.Black;
                    StuSeries2.SymbolStyle.Shape = SymbolShapeEnum.None;
                    StuSeries2.SymbolStyle.Color = Color.DarkRed;
                    StuSeries2.X.CopyDataIn(ax);
                    StuSeries2.Y.CopyDataIn(ay2);
                    StuSeries2 = null;
                }
            }

            double[] zoney = new double[] { dy1[0], dy1[2] };
            // Add and show the alarm zones
            AlarmZonesCollection zones = c1Chart1.ChartArea.PlotArea.AlarmZones;
            zones.Clear();
            for (int i = 0; i < 3; i++)
            {
                AlarmZone zone = zones.AddNewZone();

                zone.Name = dyName[i];
                zone.BackColor = dyColor[i];

                if (i == 0)
                    zone.UpperExtent = c1Chart1.ChartArea.AxisY.Max;
                else
                    zone.UpperExtent = zones[i - 1].LowerExtent;

                if (i == 2)
                    zone.LowerExtent = c1Chart1.ChartArea.AxisY.Min;
                else
                    zone.LowerExtent = zoney[i];

                zone.Visible = true;
            }

            double[] zoney2 = new double[] { dy2[0], dy2[2] };
            // Add and show the alarm zones
             zones = c1Chart2.ChartArea.PlotArea.AlarmZones;
            zones.Clear();
            for (int i = 0; i < 3; i++)
            {
                AlarmZone zone = zones.AddNewZone();

                zone.Name = dyName[i];
                zone.BackColor = dyColor[i];

                if (i == 0)
                    zone.UpperExtent = c1Chart2.ChartArea.AxisY.Max;
                else
                    zone.UpperExtent = zones[i - 1].LowerExtent;

                if (i == 2)
                    zone.LowerExtent = c1Chart2.ChartArea.AxisY.Min;
                else
                    zone.LowerExtent = zoney2[i];

                zone.Visible = true;
            }

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

        private double[] calJicha(double[] source)
        {
            int length=source.Length;
            if (length < 2)
                return null;
            double[] result = new double[length - 1];
            for (int i = 0; i < length - 1; i++)
            {
                result[i] = Math.Abs(source[i + 1] - source[i]);
            }
            return result;
        }

    }
}