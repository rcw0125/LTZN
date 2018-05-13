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
    public partial class Mn趋势 : Form
    {
        public Mn趋势()
        {
            InitializeComponent();
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
          
            ChartData data = c1Chart1.ChartGroups[0].ChartData;
            ChartDataSeriesCollection series = data.SeriesList;

          

            double[] dy1;
            //double[] dy2;
            double max;
            double min;
           
         
           LegendEnviroment.loadCanShuSetting(gaolu, "Mn", out dy1, out max, out min);
            c1Chart1.ChartArea.AxisY.Max = max;
            c1Chart1.ChartArea.AxisY.Min = min;

       

            Color[] dyColor = new Color[] { Color.FromArgb(255, 192, 192),  Color.FromArgb(255, 255, 192), Color.FromArgb(192, 255, 192) };
            string[] dyName = new string[] { "过大", "正常", "过小" };


            //plot the student scores
            ChartDataSeries StuSeries = series.AddNewSeries();
            StuSeries.Label = "TiO2(炉温)趋势";
            StuSeries.LineStyle.Pattern = LinePatternEnum.Solid;
            StuSeries.LineStyle.Color = Color.DarkBlue;
            StuSeries.SymbolStyle.Shape = SymbolShapeEnum.Star;
            StuSeries.SymbolStyle.Color = Color.DarkRed;

       

            if (r1.RadioChecked)
            {
                c1Chart1.ChartArea.AxisX.AnnoFormat = FormatEnum.DateManual;
                c1Chart1.ChartArea.AxisX.AnnoFormatString = "MM/dd HH:mm";
                c1Chart1.ChartArea.AxisX.AnnotationRotation = 0;
                
                DateTime[] sT;
                double[] s;
              
                LegendEnviroment.loadTiO2(gaolu,"FEMN", sjBegin.Value, sjEnd.Value, out sT, out s);
                StuSeries.X.CopyDataIn(sT);
                StuSeries.Y.CopyDataIn(s);
                StuSeries = null;
              

                DateTime[] ax = new DateTime[] { sjBegin.Value, sjEnd.Value };

                for (int i = 0; i < 2; i++)
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

                }

            }
            else
            {
                c1Chart1.ChartArea.AxisX.AnnoFormat = FormatEnum.NumericGeneral;
                c1Chart1.ChartArea.AxisX.AnnotationRotation = 90;
               
                long[] sLuci;
                double[] s;
                LegendEnviroment.loadTiO2(gaolu,"FEMN", toolStrip_Luci1.Text, toolStrip_Luci2.Text, out sLuci, out s);
                StuSeries.X.CopyDataIn(sLuci);
                StuSeries.Y.CopyDataIn(s);
                StuSeries = null;

             
                long[] ax = new long[] { Convert.ToInt64(toolStrip_Luci1.Text), Convert.ToInt64(toolStrip_Luci2.Text) };

                for (int i = 0; i < 2; i++)
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

                }
            }

            double[] zoney = new double[] { dy1[0], dy1[1] };
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

        private void Mn趋势_Load(object sender, EventArgs e)
        {
              splitContainer1.SplitterDistance = 200;
            this.sjBegin.Value = DateTime.Today.AddDays(-2);
            this.sjEnd.Value = DateTime.Now;
        
        }

    
    }
}