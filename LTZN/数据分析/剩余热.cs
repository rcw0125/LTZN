using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using C1.Win.C1Chart;
using System.Data.OracleClient;

namespace LTZN.数据分析
{
    public partial class 剩余热 : Form
    {
        public 剩余热()
        {
            InitializeComponent();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (sjEnd.Value < sjBegin.Value)
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

            DateTime[] shengyureT;
            double[] shengyuere;
            DateTime[] yurebiliT;
            double[] yurebili;
            DateTime[] siT;
            double[] si;

            LegendEnviroment.loadSYR(gaolu, sjBegin.Value, sjEnd.Value,out shengyureT,out shengyuere,out yurebiliT,out yurebili,out siT,out si);

            // Add the data
            ChartData data = c1Chart1.ChartGroups[0].ChartData;
            ChartDataSeriesCollection series = data.SeriesList;

            c1Chart1.ChartArea.AxisX.AnnoFormat = FormatEnum.DateManual;
            c1Chart1.ChartArea.AxisX.AnnoFormatString = "MM/dd HH:mm";
            c1Chart1.ChartArea.AxisX.AnnotationRotation = 0;

            //plot the student scores
            ChartDataSeries StuSeries = series.AddNewSeries();
            StuSeries.Label = "剩余热";
            StuSeries.LineStyle.Pattern = LinePatternEnum.Solid;
            StuSeries.LineStyle.Color = Color.DarkGoldenrod;
            StuSeries.SymbolStyle.Shape = SymbolShapeEnum.None;
            StuSeries.X.CopyDataIn(shengyureT);
            StuSeries.Y.CopyDataIn(shengyuere);
            StuSeries = null;

            //plot the student scores
            StuSeries = series.AddNewSeries();
            StuSeries.Label = "余热比例";
            StuSeries.LineStyle.Pattern = LinePatternEnum.Solid;
            StuSeries.LineStyle.Color = Color.DarkGray;
            StuSeries.SymbolStyle.Shape = SymbolShapeEnum.None;
            StuSeries.X.CopyDataIn(yurebiliT);
            StuSeries.Y.CopyDataIn(yurebili);
            StuSeries = null;

            //plot the student scores
             StuSeries = series.AddNewSeries();
            StuSeries.Label = "炉温(Si)";
            StuSeries.LineStyle.Pattern = LinePatternEnum.Solid;
            StuSeries.LineStyle.Color = Color.Red;
            StuSeries.SymbolStyle.Shape = SymbolShapeEnum.None;
            StuSeries.X.CopyDataIn(siT);
            StuSeries.Y.CopyDataIn(si);
            StuSeries = null;

        }

        private void 剩余热_Load(object sender, EventArgs e)
        {
            this.sjBegin.Value = DateTime.Today.AddDays(-2);
            this.sjEnd.Value = DateTime.Now;
        }
    }
}