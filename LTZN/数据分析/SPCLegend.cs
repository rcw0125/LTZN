using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using C1.Win.C1Chart;
using System.Data.OracleClient;

namespace LTZN.数据分析
{
    public partial class SPCLegend : UserControl
    {
        public SPCLegend()
        {
            InitializeComponent();
        }

        private void SPCLegend_Load(object sender, EventArgs e)
        {
            c1Chart1.ChartGroups[0].ChartData.SeriesList.Clear();
            this.sjBegin.Value = DateTime.Now.AddDays(-1);
            this.sjEnd.Value = DateTime.Now;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            c1Chart1.ChartGroups[0].ChartData.SeriesList.Clear();

            List<DateTime> dts = new List<DateTime>();
            List<double> vs = new List<double>();
            DateTime dt1=DateTime.Now.AddDays(-1);
            DateTime dt2=DateTime.Now;
           
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
            }
            
            OracleConnection conn = new OracleConnection(Properties.Settings.Default.ltznConnectionString);
            conn.Open();
            string sql = "select zdsj,FESi from ddluci where zdsj>=:sjBegin and zdsj<=:sjEnd and gaolu=:gaolu";
            OracleCommand siCmd = new OracleCommand(sql, conn);
            siCmd.Parameters.Add(":sjBegin", OracleType.DateTime).Value = this.sjBegin.Value;
            siCmd.Parameters.Add(":sjEnd", OracleType.DateTime).Value = this.sjEnd.Value;
            siCmd.Parameters.Add(":gaolu", OracleType.Int32).Value = gaolu;
            OracleDataReader dr = siCmd.ExecuteReader();

            while (dr.Read())
            {
                dts.Add(dr.GetDateTime(0));
                vs.Add(dr.GetDouble(1));
            }
            dr.Close();
            conn.Close();

            c1Chart1.Header.Text = this.toolStrip_gaolu.Text + "  " + this.sjBegin.Text + "--" + this.sjEnd.Text + "Si(炉温)趋势";
            
            // Add the data
            ChartData data = c1Chart1.ChartGroups[0].ChartData;
            ChartDataSeriesCollection series = data.SeriesList;

            //plot the student scores
            ChartDataSeries StuSeries = series.AddNewSeries();
            StuSeries.Label = "Si(炉温)趋势";
            StuSeries.LineStyle.Pattern = LinePatternEnum.Solid;
            StuSeries.LineStyle.Color = Color.DarkBlue;
            StuSeries.SymbolStyle.Shape = SymbolShapeEnum.Star;
            StuSeries.SymbolStyle.Color = Color.DarkRed;
            StuSeries.X.CopyDataIn(dts.ToArray());
            StuSeries.Y.CopyDataIn(vs.ToArray());
            StuSeries = null;


            DateTime[] ax = new DateTime[] { dts[0], dts[dts.Count - 1] };
            double[] ay1 = new double[] { 0.7, 0.7 };
            double[] ay2 = new double[] { 0.4, 0.4 };

            StuSeries = series.AddNewSeries();
            StuSeries.Label = "最大值";
            StuSeries.LineStyle.Pattern = LinePatternEnum.Solid;
            StuSeries.LineStyle.Color = Color.Black;
            StuSeries.SymbolStyle.Shape = SymbolShapeEnum.None;
            StuSeries.SymbolStyle.Color = Color.DarkRed;
            StuSeries.X.CopyDataIn(ax);
            StuSeries.Y.CopyDataIn(ay1);
            StuSeries = null;


            StuSeries = series.AddNewSeries();
            StuSeries.Label = "最小值";
            StuSeries.LineStyle.Pattern = LinePatternEnum.Solid;
            StuSeries.LineStyle.Color = Color.Black;
            StuSeries.SymbolStyle.Shape = SymbolShapeEnum.None;
            StuSeries.SymbolStyle.Color = Color.DarkRed;
            StuSeries.X.CopyDataIn(ax);
            StuSeries.Y.CopyDataIn(ay2);
            StuSeries = null;

            
            // Add and show the alarm zones
            AlarmZonesCollection zones = c1Chart1.ChartArea.PlotArea.AlarmZones;
            zones.Clear();

            AlarmZone zone1 = zones.AddNewZone();
            AlarmZone zone2 = zones.AddNewZone();
            AlarmZone zone3 = zones.AddNewZone();

            zone1.Name = "大值区";
            zone1.BackColor = Color.Aqua;
            zone1.UpperExtent = c1Chart1.ChartArea.AxisY.Max;
            zone1.LowerExtent = 0.7;
            zone1.Visible = true;

            
            zone2.Name = "正常";
            zone2.BackColor = Color.Beige;
            zone2.UpperExtent = 0.7;
            zone2.LowerExtent = 0.4;
            zone2.Visible = true;

            
            zone3.Name = "小值区";
            zone3.BackColor = Color.Aqua;
            zone3.UpperExtent = 0.4;
            zone3.LowerExtent = c1Chart1.ChartArea.AxisY.Min;
            zone3.Visible = true;

        }
    }
}
