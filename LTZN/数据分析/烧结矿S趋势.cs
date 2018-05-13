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
    public partial class 烧结矿S趋势 : Form
    {
        public 烧结矿S趋势()
        {
            InitializeComponent();
          
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if  (sjEnd.Value < sjBegin.Value)
            {
                MessageBox.Show("结束时间应该比开始时间大！");
                return;
            }
            int gaolu = 0;
            switch (this.toolStrip_gaolu.Text)
            {

                case "1#大烧":
                    gaolu = 7;
                    break;
                case "2#大烧":
                    gaolu = 8;
                    break;



            }

            c1Chart1.ChartGroups[0].ChartData.SeriesList.Clear();
          
            ChartData data = c1Chart1.ChartGroups[0].ChartData;
            ChartDataSeriesCollection series = data.SeriesList;

          

            double[] dy1;
            //double[] dy2;
            double max;
            double min;
           
         
           LegendEnviroment.loadCanShuSetting(gaolu, "烧结S", out dy1, out max, out min);
            c1Chart1.ChartArea.AxisY.Max = max;
            c1Chart1.ChartArea.AxisY.Min = min;

       

            Color[] dyColor = new Color[] { Color.FromArgb(255, 192, 192),  Color.FromArgb(255, 255, 192), Color.FromArgb(192, 255, 192) };
            string[] dyName = new string[] { "过大", "正常", "过小" };


            //plot the student scores
            ChartDataSeries StuSeries = series.AddNewSeries();
            StuSeries.Label = "烧结矿S趋势";
            StuSeries.LineStyle.Pattern = LinePatternEnum.Solid;
            StuSeries.LineStyle.Color = Color.DarkBlue;
            StuSeries.SymbolStyle.Shape = SymbolShapeEnum.Star;
            StuSeries.SymbolStyle.Color = Color.DarkRed;

                c1Chart1.ChartArea.AxisX.AnnoFormat = FormatEnum.DateManual;
                c1Chart1.ChartArea.AxisX.AnnoFormatString = "MM/dd HH:mm";
                c1Chart1.ChartArea.AxisX.AnnotationRotation = 0;
                
                DateTime[] sT;
                double[] s;
              
                LegendEnviroment.loadshaojie(gaolu,"S", sjBegin.Value, sjEnd.Value, out sT, out s);
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

        //private void r1_CheckedChanged(object sender, EventArgs e)
        //{
        //    if (r1.RadioChecked)
        //    {
        //        sjBegin.Visible = true;
        //        sjEnd.Visible = true;
        //        toolStrip_Luci1.Visible = false;
        //        toolStrip_Luci2.Visible = false;
        //        toolStripLabel1.Text = "开始时间";
        //        toolStripLabel2.Text = "结束时间";
        //    }
        //    else
        //    {
        //        sjBegin.Visible = false;
        //        sjEnd.Visible = false;
        //        toolStrip_Luci1.Visible = true;
        //        toolStrip_Luci2.Visible = true;
        //        toolStripLabel1.Text = "开始炉次";
        //        toolStripLabel2.Text = "结束炉次";
        //    }
        //}



        private void 烧结矿S趋势_Load(object sender, EventArgs e)
        {
              splitContainer1.SplitterDistance = 200;
            this.sjBegin.Value = DateTime.Today.AddDays(-2);
            this.sjEnd.Value = DateTime.Now;
        
        }

    
    }
}