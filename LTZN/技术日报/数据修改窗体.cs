using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;

namespace LTZN.技术日报
{
    public partial class 数据修改窗体 : Form
    {
        private bool canUpdate = false;

        public 数据修改窗体()
        {
            InitializeComponent();
            LtznUserManager.instance.RegisterHandler(this, 数据修改窗体_UserChange);
        }

        private void 数据修改窗体_Load(object sender, EventArgs e)
        {
            this.stripRiqi.Value = DateTime.Today;
        }

        private void getData(DateTime riqi)
        {
            if (canUpdate)
            {
                this.技术日报BindingSource.CurrencyManager.EndCurrentEdit();
                try
                {
                    this.技术日报TableAdapter1.Update(this.技术日报数据集1.技术日报);
                }
                catch (Exception e)
                {
                    Console.Write(e);
                }
            }
            this.技术日报TableAdapter1.FillByRq(this.技术日报数据集1.技术日报,riqi);
            if (this.技术日报数据集1.技术日报.Count != 12)
            {
                if (this.技术日报数据集1.技术日报.Count > 0)
                {
                   // foreach (技术日报数据集.技术日报Row r in this.技术日报数据集1.技术日报)
                    //{
                      //  r.Delete();
                   // }
                    this.技术日报数据集1.技术日报.Clear();
                    //  this.技术日报TableAdapter1.Update(this.技术日报数据集1.技术日报);
                    if (this.技术日报TableAdapter1.Connection.State != ConnectionState.Open)
                        this.技术日报TableAdapter1.Connection.Open();
                    OracleCommand oraCmd = new OracleCommand("delete from \"技术日报\" Where \"P日期\"=TO_DATE('" + riqi.ToString("yyyy-MM-dd") + "','YYYY-MM-DD')", this.技术日报TableAdapter1.Connection);
                    oraCmd.ExecuteNonQuery();
                }
                this.技术日报数据集1.技术日报.Add技术日报Row(riqi, "1#", "本日");
                this.技术日报数据集1.技术日报.Add技术日报Row(riqi, "1#", "累计");
                this.技术日报数据集1.技术日报.Add技术日报Row(riqi, "2#", "本日");
                this.技术日报数据集1.技术日报.Add技术日报Row(riqi, "2#", "累计");
                this.技术日报数据集1.技术日报.Add技术日报Row(riqi, "3#", "本日");
                this.技术日报数据集1.技术日报.Add技术日报Row(riqi, "3#", "累计");
                this.技术日报数据集1.技术日报.Add技术日报Row(riqi, "4#", "本日");
                this.技术日报数据集1.技术日报.Add技术日报Row(riqi, "4#", "累计");
                this.技术日报数据集1.技术日报.Add技术日报Row(riqi, "5#", "本日");
                this.技术日报数据集1.技术日报.Add技术日报Row(riqi, "5#", "累计");
                this.技术日报数据集1.技术日报.Add技术日报Row(riqi, "Q", "本日");
                this.技术日报数据集1.技术日报.Add技术日报Row(riqi, "Q", "累计");
            }
           // this.技术日报BindingSource.DataSource = this.技术日报数据集1.技术日报;
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            this.技术日报BindingSource.CurrencyManager.EndCurrentEdit();
            报表显示窗体 f = new 报表显示窗体();
            f.WindowState = FormWindowState.Maximized;
            f.setDataSource(this.技术日报数据集1.技术日报.技术日报内容);
            f.Show();
            this.Cursor = Cursors.Default;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void stripRiqi_ValueChanged(object sender, EventArgs e)
        {
            getData(this.stripRiqi.Value.Date);
        }

        private void 数据修改窗体_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (canUpdate)
            {
                this.技术日报BindingSource.CurrencyManager.EndCurrentEdit();
                this.技术日报TableAdapter1.Update(this.技术日报数据集1.技术日报);
            }
        }


        private void tiqu_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            技术日报内容 r = this.技术日报数据集1.技术日报.技术日报内容;
            生铁 st = new 生铁();
            st.getDataBy(this.stripRiqi.Value.Date);

            原料消耗 ylxh = new 原料消耗();
            ylxh.getDataBy(this.stripRiqi.Value.Date);

            for (int gaolu = 1; gaolu <= 6; gaolu++)
            {
                技术日报内容项 x = null;
                if (gaolu == 6)
                    x = r.Get技术日报内容项("Q", "本日");
                else
                    x = r.Get技术日报内容项(gaolu + "#", "本日");

                x.P03合格铁 = st.合格铁[gaolu - 1];
                x.P04炼钢铁 = st.炼钢铁[gaolu - 1];
                x.P05铸造铁 = st.铸造铁[gaolu - 1];
                x.P06号外铁 = st.号外铁[gaolu - 1];
                x.P07合格率 = st.合格率[gaolu - 1];
                x.P08高炉利用系数 = st.高炉利用系数[gaolu - 1];
                x.P09高炉实物系数 = st.高炉实物系数[gaolu - 1];
                x.P10原料矿合计总耗 = ylxh.总耗[gaolu - 1];
                x.P11原料矿合计单耗 = ylxh.总耗[gaolu - 1] * 1000 / st.合格铁[gaolu - 1];
                x.P11原料矿合计单耗 = Convert.ToDouble(x.P11原料矿合计单耗.Value.ToString("###0"));
                x.P12原料矿机烧 = ylxh.机烧[gaolu - 1];
                x.P13原料矿竖炉球 = ylxh.竖球[gaolu - 1];
                x.P14原料矿CT = 0;
                x.P15原料矿其它熟料 = ylxh.熟料[gaolu - 1];
                x.P16原料矿本溪矿 = ylxh.本溪矿[gaolu - 1];
                x.P17原料矿其它生料 = ylxh.生料[gaolu - 1];
                x.P18废铁总耗 = ylxh.废铁[gaolu - 1];
                x.P19废铁单耗 = ylxh.废铁[gaolu - 1] * 1000 / st.合格铁[gaolu - 1];
                x.P19废铁单耗 = Convert.ToDouble(x.P19废铁单耗.Value.ToString("###0"));
                x.P20回收率 = (st.全铁产量[gaolu - 1] - 0.8 * ylxh.废铁[gaolu - 1]) * 100 / ylxh.总耗[gaolu - 1];
                x.P20回收率 = Convert.ToDouble(x.P20回收率.Value.ToString("##0.00"));
                x.P21熟料比 = ylxh.熟料比[gaolu - 1];
                x.P21熟料比 = Convert.ToDouble(x.P21熟料比.Value.ToString("##0.00"));
                x.P22平均风温 = ylxh.平均风温[gaolu - 1];
                x.P22平均风温 = Convert.ToDouble(x.P22平均风温.Value.ToString("###0"));
                x.P23炉顶温度 = ylxh.炉顶温度[gaolu - 1];
                x.P23炉顶温度 = Convert.ToDouble(x.P23炉顶温度.Value.ToString("###0"));
                x.P24热风压力 = ylxh.热风压力[gaolu - 1];
                x.P24热风压力 = Convert.ToDouble(x.P24热风压力.Value.ToString("###0"));
                x.P25炉顶压力 = ylxh.炉顶压力[gaolu - 1];
                x.P25炉顶压力 = Convert.ToDouble(x.P25炉顶压力.Value.ToString("###0"));
                x.P26富氧率 = ylxh.富氧率[gaolu - 1];
                x.P26富氧率 = Convert.ToDouble(x.P26富氧率.Value.ToString("##0.00"));
                x.P27入炉焦炭总耗 = ylxh.入炉焦炭[gaolu - 1];
                x.P28入炉焦炭单耗 = ylxh.入炉焦炭[gaolu - 1] * 1000 / st.合格铁[gaolu - 1];
                x.P28入炉焦炭单耗 = Convert.ToDouble(x.P28入炉焦炭单耗.Value.ToString("###0"));
                x.P29煤粉总耗 = ylxh.煤粉[gaolu - 1];
                x.P30煤粉单耗 = ylxh.煤粉[gaolu - 1] * 1000 / st.合格铁[gaolu - 1];
                x.P30煤粉单耗 = Convert.ToDouble(x.P30煤粉单耗.Value.ToString("###0"));
                x.P31焦丁总耗 = ylxh.焦丁[gaolu - 1];
                x.P32焦丁单耗 = ylxh.焦丁[gaolu - 1] * 1000 / st.合格铁[gaolu - 1];
                x.P32焦丁单耗 = Convert.ToDouble(x.P32焦丁单耗.Value.ToString("###0"));
                x.P33综合焦炭总耗 = ylxh.综合焦炭[gaolu - 1];
                x.P34综合焦炭单耗 = ylxh.综合焦炭[gaolu - 1] * 1000 / st.合格铁[gaolu - 1];
                x.P34综合焦炭单耗 = Convert.ToDouble(x.P34综合焦炭单耗.Value.ToString("###0"));
                x.P35综合折算焦比 = st.折算产量[gaolu - 1] == 0 ? 0 : ylxh.综合焦炭[gaolu - 1] * 1000 / st.折算产量[gaolu - 1];
                x.P35综合折算焦比 = Convert.ToDouble(x.P35综合折算焦比.Value.ToString("##0"));
                x.P36冶炼强度 = ylxh.冶炼强度[gaolu - 1];
                x.P36冶炼强度 = Convert.ToDouble(x.P36冶炼强度.Value.ToString("##0.00"));
                x.P37焦炭负荷 = ylxh.焦炭负荷[gaolu - 1];
                x.P37焦炭负荷 = Convert.ToDouble(x.P37焦炭负荷.Value.ToString("##0.00"));
                x.P38干毛焦 = ylxh.干毛焦[gaolu - 1];
                x.P39炼钢铁SI = st.炼钢铁Si[gaolu - 1];
                x.P39炼钢铁SI = Convert.ToDouble(x.P39炼钢铁SI.Value.ToString("##0.00"));
                x.P40炼钢铁MN = st.炼钢铁Mn[gaolu - 1];
                x.P40炼钢铁MN = Convert.ToDouble(x.P40炼钢铁MN.Value.ToString("##0.00"));
                x.P41炼钢铁P = st.炼钢铁P[gaolu - 1];
                x.P41炼钢铁P = Convert.ToDouble(x.P41炼钢铁P.Value.ToString("##0.000"));
                x.P42炼钢铁S = st.炼钢铁S[gaolu - 1];
                x.P42炼钢铁S = Convert.ToDouble(x.P42炼钢铁S.Value.ToString("##0.000"));
                x.P43铸造铁SI = st.铸造铁Si[gaolu - 1];
                x.P43铸造铁SI = Convert.ToDouble(x.P43铸造铁SI.Value.ToString("##0.00"));
                x.P44铸造铁MN = st.铸造铁Mn[gaolu - 1];
                x.P44铸造铁MN = Convert.ToDouble(x.P44铸造铁MN.Value.ToString("##0.00"));
                x.P45铸造铁P = st.铸造铁P[gaolu - 1];
                x.P45铸造铁P = Convert.ToDouble(x.P45铸造铁P.Value.ToString("##0.00"));
                x.P46铸造铁S = st.铸造铁S[gaolu - 1];
                x.P46铸造铁S = Convert.ToDouble(x.P46铸造铁S.Value.ToString("##0.00"));
                x.P47炉渣碱度 = st.炉渣碱度[gaolu - 1];
                x.P47炉渣碱度 = Convert.ToDouble(x.P47炉渣碱度.Value.ToString("##0.00"));
                x.P48休风情况 = ylxh.休风情况[gaolu - 1];
                x.P49慢风 = ylxh.慢风[gaolu - 1];
                x.P50坐料次数 = ylxh.坐料次数[gaolu - 1];
                x.P51悬料次数 = ylxh.悬料次数[gaolu - 1];
                x.P52崩料次数 = ylxh.崩料次数[gaolu - 1];
                x.P53风口损坏数大 = ylxh.风口损坏数大[gaolu - 1];
                x.P54风口损坏数中 = ylxh.风口损坏数中[gaolu - 1];
                x.P55风口损坏数小 = ylxh.风口损坏数小[gaolu - 1];
                x.P56渣口损坏数大 = ylxh.渣口损坏数大[gaolu - 1];
                x.P57渣口损坏数中 = ylxh.渣口损坏数中[gaolu - 1];
                x.P58渣口损坏数小 = ylxh.渣口损坏数小[gaolu - 1];
                x.P59本厂铸造SI大 = st.本厂铸造SI大[gaolu - 1];
                x.P60本厂铸造SI小 = st.本厂铸造SI小[gaolu - 1];
                x.P61送炼钢厂SI大 = st.送炼钢厂SI大[gaolu - 1];
                x.P62送炼钢厂SI中 = 0;
                x.P63送炼钢厂SI小 = st.送炼钢厂SI小[gaolu - 1];
                x.P64折算产量 = st.折算产量[gaolu - 1];
                x.P65工艺称焦比 = ylxh.工艺称焦炭[gaolu - 1] * ylxh.干毛焦[gaolu - 1] * 1000 / (ylxh.自产湿焦[gaolu - 1] + ylxh.落地湿焦[gaolu - 1]) / (st.全铁产量[gaolu - 1]);
                x.P生成标志 = 1;
                for (int i = 0; i < x.p.Length; i++)
                {
                    x.p[i] = true;
                }

                技术日报内容项 xl = null;
                if (gaolu == 6)
                    xl = r.Get技术日报内容项("Q", "累计");
                else
                    xl = r.Get技术日报内容项(gaolu + "#", "累计");

                xl.P03合格铁 = st.累计合格铁[gaolu - 1];
                xl.P04炼钢铁 = st.累计炼钢铁[gaolu - 1];
                xl.P05铸造铁 = st.累计铸造铁[gaolu - 1];
                xl.P06号外铁 = st.累计号外铁[gaolu - 1];
                xl.P07合格率 = st.累计合格率[gaolu - 1];
                xl.P08高炉利用系数 = st.累计高炉利用系数[gaolu - 1];
                xl.P09高炉实物系数 = st.累计高炉实物系数[gaolu - 1];
                xl.P10原料矿合计总耗 = ylxh.累计总耗[gaolu - 1];
                xl.P11原料矿合计单耗 = ylxh.累计总耗[gaolu - 1] * 1000 / st.累计合格铁[gaolu - 1];
                xl.P11原料矿合计单耗 = Convert.ToDouble(xl.P11原料矿合计单耗.Value.ToString("###0"));
                xl.P12原料矿机烧 = ylxh.累计机烧[gaolu - 1];
                xl.P13原料矿竖炉球 = ylxh.累计竖球[gaolu - 1];
                xl.P14原料矿CT = 0;
                xl.P15原料矿其它熟料 = ylxh.累计熟料[gaolu - 1];
                xl.P16原料矿本溪矿 = ylxh.累计本溪矿[gaolu - 1];
                xl.P17原料矿其它生料 = ylxh.累计生料[gaolu - 1];
                xl.P18废铁总耗 = ylxh.累计废铁[gaolu - 1];
                xl.P19废铁单耗 = ylxh.累计废铁[gaolu - 1] * 1000 / st.累计合格铁[gaolu - 1];
                xl.P19废铁单耗 = Convert.ToDouble(xl.P19废铁单耗.Value.ToString("###0"));
                xl.P20回收率 = (st.累计全铁产量[gaolu - 1] - 0.8 * ylxh.累计废铁[gaolu - 1]) * 100 / ylxh.累计总耗[gaolu - 1];
                xl.P20回收率 = Convert.ToDouble(xl.P20回收率.Value.ToString("##0.00"));
                xl.P21熟料比 = ylxh.累计熟料比[gaolu - 1];
                xl.P21熟料比 = Convert.ToDouble(xl.P21熟料比.Value.ToString("##0.00"));
                xl.P22平均风温 = ylxh.累计平均风温[gaolu - 1];
                xl.P22平均风温 = Convert.ToDouble(xl.P22平均风温.Value.ToString("###0"));
                xl.P23炉顶温度 = ylxh.累计炉顶温度[gaolu - 1];
                xl.P23炉顶温度 = Convert.ToDouble(xl.P23炉顶温度.Value.ToString("###0"));
                xl.P24热风压力 = ylxh.累计热风压力[gaolu - 1];
                xl.P24热风压力 = Convert.ToDouble(xl.P24热风压力.Value.ToString("###0"));
                xl.P25炉顶压力 = ylxh.累计炉顶压力[gaolu - 1];
                xl.P25炉顶压力 = Convert.ToDouble(xl.P25炉顶压力.Value.ToString("###0"));
                xl.P26富氧率 = ylxh.累计富氧率[gaolu - 1];
                xl.P26富氧率 = Convert.ToDouble(xl.P26富氧率.Value.ToString("##0.00"));
                xl.P27入炉焦炭总耗 = ylxh.累计入炉焦炭[gaolu - 1];
                xl.P28入炉焦炭单耗 = ylxh.累计入炉焦炭[gaolu - 1] * 1000 / st.累计合格铁[gaolu - 1];
                xl.P28入炉焦炭单耗 = Convert.ToDouble(xl.P28入炉焦炭单耗.Value.ToString("###0"));
                xl.P29煤粉总耗 = ylxh.累计煤粉[gaolu - 1];
                xl.P30煤粉单耗 = ylxh.累计煤粉[gaolu - 1] * 1000 / st.累计合格铁[gaolu - 1];
                xl.P30煤粉单耗 = Convert.ToDouble(xl.P30煤粉单耗.Value.ToString("###0"));
                xl.P31焦丁总耗 = ylxh.累计焦丁[gaolu - 1];
                xl.P32焦丁单耗 = ylxh.累计焦丁[gaolu - 1] * 1000 / st.累计合格铁[gaolu - 1];
                xl.P32焦丁单耗 = Convert.ToDouble(xl.P32焦丁单耗.Value.ToString("###0"));
                xl.P33综合焦炭总耗 = ylxh.累计综合焦炭[gaolu - 1];
                xl.P34综合焦炭单耗 = ylxh.累计综合焦炭[gaolu - 1] * 1000 / st.累计合格铁[gaolu - 1];
                xl.P34综合焦炭单耗 = Convert.ToDouble(xl.P34综合焦炭单耗.Value.ToString("###0"));
                xl.P35综合折算焦比 = st.累计折算产量[gaolu - 1] == 0 ? 0 : ylxh.累计综合焦炭[gaolu - 1] * 1000 / st.累计折算产量[gaolu - 1];
                xl.P35综合折算焦比 = Convert.ToDouble(xl.P35综合折算焦比.Value.ToString("##0"));
                xl.P36冶炼强度 = ylxh.累计冶炼强度[gaolu - 1];
                xl.P36冶炼强度 = Convert.ToDouble(xl.P36冶炼强度.Value.ToString("##0.00"));
                xl.P37焦炭负荷 = ylxh.累计焦炭负荷[gaolu - 1];
                xl.P37焦炭负荷 = Convert.ToDouble(xl.P37焦炭负荷.Value.ToString("##0.00"));
                xl.P38干毛焦 = ylxh.累计干毛焦[gaolu - 1];
                xl.P39炼钢铁SI = st.累计炼钢铁Si[gaolu - 1];
                xl.P39炼钢铁SI = Convert.ToDouble(xl.P39炼钢铁SI.Value.ToString("##0.00"));
                xl.P40炼钢铁MN = st.累计炼钢铁Mn[gaolu - 1];
                xl.P40炼钢铁MN = Convert.ToDouble(xl.P40炼钢铁MN.Value.ToString("##0.00"));
                xl.P41炼钢铁P = st.累计炼钢铁P[gaolu - 1];
                xl.P41炼钢铁P = Convert.ToDouble(xl.P41炼钢铁P.Value.ToString("##0.000"));
                xl.P42炼钢铁S = st.累计炼钢铁S[gaolu - 1];
                xl.P42炼钢铁S = Convert.ToDouble(xl.P42炼钢铁S.Value.ToString("##0.000"));
                xl.P43铸造铁SI = st.累计铸造铁Si[gaolu - 1];
                xl.P43铸造铁SI = Convert.ToDouble(xl.P43铸造铁SI.Value.ToString("##0.00"));
                xl.P44铸造铁MN = st.累计铸造铁Mn[gaolu - 1];
                xl.P44铸造铁MN = Convert.ToDouble(xl.P44铸造铁MN.Value.ToString("##0.00"));
                xl.P45铸造铁P = st.累计铸造铁P[gaolu - 1];
                xl.P45铸造铁P = Convert.ToDouble(xl.P45铸造铁P.Value.ToString("##0.00"));
                xl.P46铸造铁S = st.累计铸造铁S[gaolu - 1];
                xl.P46铸造铁S = Convert.ToDouble(xl.P46铸造铁S.Value.ToString("##0.00"));
                xl.P47炉渣碱度 = st.累计炉渣碱度[gaolu - 1];
                xl.P47炉渣碱度 = Convert.ToDouble(xl.P47炉渣碱度.Value.ToString("##0.00"));
                xl.P48休风情况 = ylxh.累计休风情况[gaolu - 1];
                xl.P49慢风 = ylxh.累计慢风[gaolu - 1];
                xl.P50坐料次数 = ylxh.累计坐料次数[gaolu - 1];
                xl.P51悬料次数 = ylxh.累计悬料次数[gaolu - 1];
                xl.P52崩料次数 = ylxh.累计崩料次数[gaolu - 1];
                xl.P53风口损坏数大 = ylxh.累计风口损坏数大[gaolu - 1];
                xl.P54风口损坏数中 = ylxh.累计风口损坏数中[gaolu - 1];
                xl.P55风口损坏数小 = ylxh.累计风口损坏数小[gaolu - 1];
                xl.P56渣口损坏数大 = ylxh.累计渣口损坏数大[gaolu - 1];
                xl.P57渣口损坏数中 = ylxh.累计渣口损坏数中[gaolu - 1];
                xl.P58渣口损坏数小 = ylxh.累计渣口损坏数小[gaolu - 1];
                xl.P59本厂铸造SI大 = st.累计本厂铸造SI大[gaolu - 1];
                xl.P60本厂铸造SI小 = st.累计本厂铸造SI小[gaolu - 1];
                xl.P61送炼钢厂SI大 = st.累计送炼钢厂SI大[gaolu - 1];
                xl.P62送炼钢厂SI中 = 0;
                xl.P63送炼钢厂SI小 = st.累计送炼钢厂SI小[gaolu - 1];
                x.P64折算产量 = st.累计折算产量[gaolu - 1];
                x.P65工艺称焦比 = Convert.ToDouble((ylxh.累计工艺称焦炭[gaolu - 1] * ylxh.累计干毛焦[gaolu - 1] * 1000 / (ylxh.累计自产湿焦[gaolu - 1] + ylxh.累计落地湿焦[gaolu - 1]) / (st.累计全铁产量[gaolu - 1])).ToString("####0"));

                for (int i = 0; i < xl.p.Length ; i++)
                {
                    xl.p[i] = true;
                }

            }
            this.技术日报数据集1.技术日报.技术日报内容 = r;
            this.Cursor = Cursors.Default;
        }

        void 数据修改窗体_UserChange(LtznUser ltznUser)
        {
            if (ltznUser!=null && ltznUser.IsInRole("统计"))
            {
                canUpdate = true;
                tiqu.Enabled = true;
            }
            else
            {
                canUpdate = false;
                tiqu.Enabled = false;
            }
        }
    }
}