using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using C1.Win.C1FlexGrid;

namespace LTZN.数据分析
{
    public partial class MaterialAnalysisForm : Form
    {
        public MaterialAnalysisForm()
        {
            InitializeComponent();
        }

        private void MaterialAnalysisForm_Load(object sender, EventArgs e)
        {
            sjBegin.Value = DateTime.Now.AddDays(-1);
            sjEnd.Value = DateTime.Now;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            Dictionary<string, string> nameMap = new Dictionary<string, string>();
            nameMap.Add("MC", "名称");
            nameMap.Add("PZ", "品种");
            nameMap.Add("HUIFEN", "灰份");
            nameMap.Add("HUIFA", "挥发性");
            nameMap.Add("SF", "水份");
            nameMap.Add("QD", "强度");
            nameMap.Add("TFE", "TFe");
            nameMap.Add("FEO", "FeO");
            nameMap.Add("SIO2", "SiO2");
            nameMap.Add("CAO", "CaO");
            nameMap.Add("MgO", "MgO");
            nameMap.Add("ALO", "Al2O3");
            nameMap.Add("TIO2", "TiO2");
            nameMap.Add("JD", "碱度");

            switch (toolStrip_fenlei.Text)
            {
                case "原料":
                    c1FlexGrid2.DataMember = "DDYL";
                    MaterialAnalysisDataSetTableAdapters.DDYLTableAdapter ylAdp = new LTZN.数据分析.MaterialAnalysisDataSetTableAdapters.DDYLTableAdapter();
                    ylAdp.Fill(this.materialAnalysisDataSet.DDYL, sjBegin.Value, sjEnd.Value);

                    break;

                case "煤粉":
                    c1FlexGrid2.DataMember = "DDMF";
                    MaterialAnalysisDataSetTableAdapters.DDMFTableAdapter mfAdp = new LTZN.数据分析.MaterialAnalysisDataSetTableAdapters.DDMFTableAdapter();
                    mfAdp.Fill(this.materialAnalysisDataSet.DDMF, sjBegin.Value, sjEnd.Value);
                    break;
                case "焦炭":
                    c1FlexGrid2.DataMember = "DDJT";
                    MaterialAnalysisDataSetTableAdapters.DDJTTableAdapter jtAdp = new LTZN.数据分析.MaterialAnalysisDataSetTableAdapters.DDJTTableAdapter();
                    jtAdp.Fill(this.materialAnalysisDataSet.DDJT, sjBegin.Value, sjEnd.Value);
                    break;

            }
            foreach (C1.Win.C1FlexGrid.Column col in c1FlexGrid2.Cols)
            {
                if (nameMap.ContainsKey(col.Name))
                {
                    col.Caption = nameMap[col.Name];
                }
                else
                    col.Caption = col.Name;
            }
            //c1FlexGrid2.Subtotal(AggregateEnum.Clear);
            //c1FlexGrid2.SubtotalPosition = SubtotalPositionEnum.BelowData;
            ////for (int i = 1; i < c1FlexGrid2.Cols.Count - 1; i++)
            ////{
            //c1FlexGrid2.Subtotal(AggregateEnum.Average, -1, -1, c1FlexGrid2.Cols.Count - 2, "平均:");
            ////}
            //CellStyle s = c1FlexGrid2.Styles[CellStyleEnum.GrandTotal];
            //s.BackColor = Color.Azure;
            //s.ForeColor = Color.Black;
        }


    }



}
