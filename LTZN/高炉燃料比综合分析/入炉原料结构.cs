using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using C1.Win.C1Chart;

namespace LTZN.高炉燃料比综合分析
{
    public partial class 入炉原料结构 : Form
    {

        DateTime[] sT;
        double[] s;
        List<string> ylmcysj = new List<string>();
        public 入炉原料结构()
        {
            InitializeComponent();
        }

        private void 入炉原料结构_Load(object sender, EventArgs e)
        {
            splitContainer1.SplitterDistance = 200;
            this.sjBegin.Value = DateTime.Today.AddDays(-2);
            this.sjEnd.Value = DateTime.Now;
        }
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (sjEnd.Value < sjBegin.Value)
            {
                MessageBox.Show("结束时间应该比开始时间大！");
                return;
            }
           
            c1Chart1.ChartGroups[0].ChartData.SeriesList.Clear();
            Color jthf = Color.FromArgb(25, 45, 86);

            // Add the data
            ChartData data = c1Chart1.ChartGroups[0].ChartData;
            ChartDataSeriesCollection series = data.SeriesList;

            ChartData data1 = c1Chart1.ChartGroups[1].ChartData;
            ChartDataSeriesCollection series1 = data1.SeriesList;
            series.Clear();
            series1.Clear();
            ylmcysj.Clear();
 
            c1Chart1.ChartArea.AxisY.Max = 100;
            c1Chart1.ChartArea.AxisY.Min = 0;
            c1Chart1.ChartArea.AxisY.Text = "百分比";
            List<string> ylmc = new List<string>();
            LegendEnviroment.loadGLYL("含铁原料",out ylmc);
            int i = 0;
            foreach(String yl in ylmc)
            {
                LegendEnviroment.loadRlylPeiBi(Convert.ToInt32(gaolu.Text), yl, sjBegin.Value, sjEnd.Value, out sT, out s);
          
                if(s.Length>1)
                {
                    ChartDataSeries StuSeries = series.AddNewSeries();
                    StuSeries.Label = yl+"配比";
                    StuSeries.LineStyle.Pattern = LinePatternEnum.Solid;
                    Random r = new Random();
                    StuSeries.LineStyle.Color = colors.getColor(i++);// Color.FromArgb(r.Next(255), r.Next(255), r.Next(255));
                    StuSeries.LineStyle.Thickness = 2;
                    StuSeries.SymbolStyle.Shape = SymbolShapeEnum.Dot;
                    StuSeries.SymbolStyle.Color = StuSeries.LineStyle.Color;
                    StuSeries.X.CopyDataIn(sT);
                    StuSeries.Y.CopyDataIn(s);
                    ylmcysj.Add(yl);
                }
            }
          
            ////plot the student scores
      
                ChartDataSeries StuSeries2 = series.AddNewSeries();
                StuSeries2.Label = "熔剂配比";
                StuSeries2.LineStyle.Pattern = LinePatternEnum.Solid;
                StuSeries2.LineStyle.Color = Color.FromArgb(218, 129, 55);
                StuSeries2.LineStyle.Thickness = 2;
                StuSeries2.SymbolStyle.Shape = SymbolShapeEnum.Dot;
                StuSeries2.SymbolStyle.Color = Color.FromArgb(218, 129, 55);
                ylmcysj.Add("熔剂");
      
                // 熔剂配比
                LegendEnviroment.loadRONGJIPeiBi(Convert.ToInt32(gaolu.Text), sjBegin.Value, sjEnd.Value, out sT, out s);
                StuSeries2.X.CopyDataIn(sT);
                StuSeries2.Y.CopyDataIn(s);

                c1Chart1.ChartArea.AxisX.AnnoFormat = FormatEnum.DateManual;
                c1Chart1.ChartArea.AxisX.AnnoFormatString = "MM/dd";
                c1Chart1.ChartArea.AxisX.UnitMajor = 1;
                c1Chart1.ChartArea.AxisX.AnnotationRotation = 0;

                c1Chart1.Legend.Visible = true;
                Area carea = c1Chart1.ChartArea;
                carea.AxisX.ValueLabels.Clear();
                double xMarker = AddChartAxisMarker(carea.AxisX);
              
        }
        private double AddChartAxisMarker(Axis ax)
        {
            // Set the annotation method to Mixed so both automatic
            // Values annotation AND ValueLabels can be used concurrently.
            ax.AnnoMethod = AnnotationMethodEnum.Mixed;

            // Set the GridMajor line pattern to solid.  This will be
            // used for the ValueLabel Gridline draw.
            ax.GridMajor.Pattern = LinePatternEnum.Solid;
            ax.OnTop = true;

            // A convenient value for the axes units.
            ax.UnitMajor = 10;

            // Create the ValueLabel.  Use an arrow, and make it big
            // and red so it cannot be missed.
            ValueLabel vl = ax.ValueLabels.AddNewLabel();
            vl.Appearance = ValueLabelAppearanceEnum.ArrowMarker;
            vl.GridLine = true;
            vl.Moveable = true;
            vl.MarkerSize = 20;
            vl.Color = Color.Red;

            // Put it in the middle of the axis.  For this to work
            // the chart must have been drawn so the axis is properly
            // scaled.  A prior GetImage() will do this.
           
            vl.DateTimeValue = sjEnd.Value.AddDays(-2);

            // Set the event so the data can be recalculated
            // each time the marker is moved.
            vl.ValueChanged += new EventHandler(MarkerMoved);	//!!VBSubst AddHandler vl.ValueChanged, AddressOf MarkerMoved

            // return the numeric location.
            return vl.NumericValue;
        }
        private void MarkerMoved(object sender, EventArgs e)
        {
            SetHistogramData();
        }
        private void SetHistogramData()
        {
           
            Area carea = c1Chart1.ChartArea;
            if (carea.AxisX.ValueLabels[0] != null)
            {
                // get the Target Coordinates
                DateTime xtarget = carea.AxisX.ValueLabels[0].DateTimeValue;
                DateTime xtime= new DateTime (xtarget.Year,xtarget.Month,xtarget.Day);
                DateTime mintime = new DateTime(sjBegin.Value.Year,sjBegin.Value.Month,sjBegin.Value.Day);

                for (int m = 0; m < c1Chart1.ChartGroups[0].ChartData.SeriesList.Count; m++)
                {
                    //// get the data point coordinates from the chart.
                    ChartDataSeries cds = c1Chart1.ChartGroups[0].ChartData.SeriesList[m];
                    PointF[] cdata = (PointF[])cds.PointData.CopyDataOut();

                    double dy = 0;
                    //// find the distance from each scatter point to the target point.
                    int n = cdata.Length;
            
                    for (int i = 0; i < n; i++)
                    {
                        double dx = cdata[i].X;
                        if (dx == xtime.ToOADate())
                        {
                            dy = cdata[i].Y;
                        }
                        
                    }
                    c1Chart1.ChartGroups[0].ChartData.SeriesList[m].Label = xtime.GetDateTimeFormats('M')[0].ToString() + ylmcysj[m].ToString() + ":" + dy.ToString("0.00");
                }
            }
        }

    }
}