using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using C1.Win.C1Chart;
using C1.Win.C1FlexGrid;

namespace LTZN.高炉燃料比综合分析
{
    public partial class 数据汇总 : Form
    {
        public 数据汇总()
        {
            InitializeComponent();
        }

        private void 数据汇总_Load(object sender, EventArgs e)
        {
            splitContainer1.SplitterDistance = 200;
            this.sjBegin.Value = DateTime.Today.AddDays(-2);
            this.sjEnd.Value = DateTime.Now;

            CellStyle s1 = c1FlexGrid_tiezha.Styles[CellStyleEnum.Subtotal0];
            s1.BackColor = Color.WhiteSmoke;
            s1.ForeColor = Color.Black;
            s1.Font = new Font(c1FlexGrid_tiezha.Font, FontStyle.Bold);
            s1.Format = "#######0.##";
            c1FlexGrid_tiezha.SubtotalPosition = SubtotalPositionEnum.BelowData;
        }
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (sjEnd.Value < sjBegin.Value)
            {
                MessageBox.Show("结束时间应该比开始时间大！");
                return;
            }
            List<DataZong> huizong = new List<DataZong>();
            Dictionary<DateTime, DataZong> dz = new Dictionary<DateTime, DataZong>();

            // 自产焦炭灰分
            DateTime[] sTjiaotanhuifen;
            double[] data;
            DateTime[] sT;
            LegendEnviroment.loadZichanj(sjBegin.Value, sjEnd.Value, out sTjiaotanhuifen, out data);
            for (int i = 0; i < sTjiaotanhuifen.Length; i++)
            {

                if (dz.ContainsKey(sTjiaotanhuifen[i]))
                {
                    dz[sTjiaotanhuifen[i]].Zichanjthf = data[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sTjiaotanhuifen[i];
                    ndz.Zichanjthf = data[i];
                    dz.Add(sTjiaotanhuifen[i], ndz);
                }
            }
            DateTime[] sTjiaotanshuifen;
            double[] datasTjiaotanshuifen;
            LegendEnviroment.loadZichanjShuiFen(sjBegin.Value, sjEnd.Value, out sTjiaotanshuifen, out datasTjiaotanshuifen);

            //自产焦炭水分
            for (int i = 0; i < sTjiaotanshuifen.Length; i++)
            {

                if (dz.ContainsKey(sTjiaotanshuifen[i]))
                {
                    dz[sTjiaotanshuifen[i]].Zichanjtsf = datasTjiaotanshuifen[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sTjiaotanshuifen[i];
                    ndz.Zichanjtsf = datasTjiaotanshuifen[i];
                    dz.Add(sTjiaotanshuifen[i], ndz);
                }
            }
            DateTime[] sTWaiGouj;
            double[] dataWaiGouj;

            LegendEnviroment.loadWaiGouj(sjBegin.Value, sjEnd.Value, out sTWaiGouj, out dataWaiGouj);

            //外购焦炭灰分分
            for (int i = 0; i < sTWaiGouj.Length; i++)
            {

                if (dz.ContainsKey(sTWaiGouj[i]))
                {
                    dz[sTWaiGouj[i]].Waigoujthf = dataWaiGouj[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sTWaiGouj[i];
                    ndz.Waigoujthf = dataWaiGouj[i];
                    dz.Add(sTWaiGouj[i], ndz);
                }
            }

            DateTime[] sTWaiGoujWaiGoujtshuifen;
            double[] dataWaiGoujWaiGoujtshuifen;


            LegendEnviroment.loadWaiGoujtshuifen(sjBegin.Value, sjEnd.Value, out sTWaiGoujWaiGoujtshuifen, out dataWaiGoujWaiGoujtshuifen);

            //外购焦炭水分
            for (int i = 0; i < sTWaiGoujWaiGoujtshuifen.Length; i++)
            {

                if (dz.ContainsKey(sTWaiGoujWaiGoujtshuifen[i]))
                {
                    dz[sTWaiGoujWaiGoujtshuifen[i]].Waigoujtsf = dataWaiGoujWaiGoujtshuifen[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sTWaiGoujWaiGoujtshuifen[i];
                    ndz.Waigoujtsf = dataWaiGoujWaiGoujtshuifen[i];
                    dz.Add(sTWaiGoujWaiGoujtshuifen[i], ndz);
                }
            }
            DateTime[] sTMeiFenh;
            double[] dataMeiFenh;

            LegendEnviroment.loadMeiFenh(sjBegin.Value, sjEnd.Value, out sTMeiFenh, out dataMeiFenh);

            //喷吹煤粉灰分
            for (int i = 0; i < sTMeiFenh.Length; i++)
            {

                if (dz.ContainsKey(sTMeiFenh[i]))
                {
                    dz[sTMeiFenh[i]].Pcmfhf = dataMeiFenh[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sTMeiFenh[i];
                    ndz.Pcmfhf = dataMeiFenh[i];
                    dz.Add(sTMeiFenh[i], ndz);
                }
            }
            DateTime[] sTMeiFenhuifa;
            double[] dataMeiFenhuifa;

            LegendEnviroment.loadMeiFenhuifa(sjBegin.Value, sjEnd.Value, out sTMeiFenhuifa, out dataMeiFenhuifa);

            //喷吹煤粉挥发份

            for (int i = 0; i < sTMeiFenhuifa.Length; i++)
            {

                if (dz.ContainsKey(sTMeiFenhuifa[i]))
                {
                    dz[sTMeiFenhuifa[i]].Pcmfhff = dataMeiFenhuifa[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sTMeiFenhuifa[i];
                    ndz.Pcmfhff = data[i];
                    dz.Add(sTMeiFenhuifa[i], ndz);
                }
            }

            DateTime[] sTMeiFenShuifen;
            double[] dataMeiFenShuifen;
            LegendEnviroment.loadMeiFenShuifen(sjBegin.Value, sjEnd.Value, out sTMeiFenShuifen, out dataMeiFenShuifen);

            //喷吹煤粉水分

            for (int i = 0; i < sTMeiFenShuifen.Length; i++)
            {

                if (dz.ContainsKey(sTMeiFenShuifen[i]))
                {
                    dz[sTMeiFenShuifen[i]].Pcmfsf = dataMeiFenShuifen[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sTMeiFenShuifen[i];
                    ndz.Pcmfsf = data[i];
                    dz.Add(sTMeiFenShuifen[i], ndz);
                }
            }
            DateTime[] sTSJKPinWei;
            double[] dataSJKPinWei;
            LegendEnviroment.loadSJKPinWei(sjBegin.Value, sjEnd.Value, out sTSJKPinWei, out dataSJKPinWei);

            //烧结矿品位

            for (int i = 0; i < sTSJKPinWei.Length; i++)
            {

                if (dz.ContainsKey(sTSJKPinWei[i]))
                {
                    dz[sTSJKPinWei[i]].Sjkpw = dataSJKPinWei[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sTSJKPinWei[i];
                    ndz.Sjkpw = dataSJKPinWei[i];
                    dz.Add(sTSJKPinWei[i], ndz);
                }
            }
            DateTime[] sTQTKPinWei;
            double[] dataQTKPinWei;
            LegendEnviroment.loadQTKPinWei(sjBegin.Value, sjEnd.Value, out sTQTKPinWei, out dataQTKPinWei);

            //球团矿品位

            for (int i = 0; i < sTQTKPinWei.Length; i++)
            {

                if (dz.ContainsKey(sTQTKPinWei[i]))
                {
                    dz[sTQTKPinWei[i]].Qtkpw = dataQTKPinWei[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sTQTKPinWei[i];
                    ndz.Qtkpw = dataQTKPinWei[i];
                    dz.Add(sTQTKPinWei[i], ndz);
                }
            }
     
            LegendEnviroment.loadAOKPinWei(sjBegin.Value, sjEnd.Value, out sT, out data);

            //澳块矿品位

            for (int i = 0; i < sT.Length; i++)
            {

                if (dz.ContainsKey(sT[i]))
                {
                    dz[sT[i]].Akpw = data[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sT[i];
                    ndz.Akpw = data[i];
                    dz.Add(sT[i], ndz);
                }
            }
            LegendEnviroment.loadQTKKPinWei(sjBegin.Value, sjEnd.Value, out sT, out data);

            //其他块矿品位

            for (int i = 0; i < sT.Length; i++)
            {

                if (dz.ContainsKey(sT[i]))
                {
                    dz[sT[i]].Qtkpw = data[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sT[i];
                    ndz.Qtkkpw = data[i];
                    dz.Add(sT[i], ndz);
                }
            }
            LegendEnviroment.loadTQPinWei(sjBegin.Value, sjEnd.Value, out sT, out data);

            //钛球品位

            for (int i = 0; i < sT.Length; i++)
            {

                if (dz.ContainsKey(sT[i]))
                {
                    dz[sT[i]].Tqpw = data[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sT[i];
                    ndz.Tqpw = data[i];
                    dz.Add(sT[i], ndz);
                }
            }
            LegendEnviroment.loadRlylPeiBi(Convert.ToInt32(gaolu.Text),"烧结矿",sjBegin.Value, sjEnd.Value, out sT, out data);

            //烧结矿配比

            for (int i = 0; i < sT.Length; i++)
            {

                if (dz.ContainsKey(sT[i]))
                {
                    dz[sT[i]].Sjkpb = data[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sT[i];
                    ndz.Sjkpb = data[i];
                    dz.Add(sT[i], ndz);
                }
            }
            LegendEnviroment.loadRlylPeiBi(Convert.ToInt32(gaolu.Text), "竖炉球", sjBegin.Value, sjEnd.Value, out sT, out data);

            //球团矿配比

            for (int i = 0; i < sT.Length; i++)
            {

                if (dz.ContainsKey(sT[i]))
                {
                    dz[sT[i]].Qhkpb = data[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sT[i];
                    ndz.Qhkpb = data[i];
                    dz.Add(sT[i], ndz);
                }
            }
            LegendEnviroment.loadRlylPeiBi(Convert.ToInt32(gaolu.Text), "澳矿", sjBegin.Value, sjEnd.Value, out sT, out data);

            //澳矿配比

            for (int i = 0; i < sT.Length; i++)
            {

                if (dz.ContainsKey(sT[i]))
                {
                    dz[sT[i]].Akpb = data[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sT[i];
                    ndz.Akpb = data[i];
                    dz.Add(sT[i], ndz);
                }
            }
            LegendEnviroment.loadRlylPeiBi(Convert.ToInt32(gaolu.Text), "块", sjBegin.Value, sjEnd.Value, out sT, out data);

            //块矿配比

            for (int i = 0; i < sT.Length; i++)
            {

                if (dz.ContainsKey(sT[i]))
                {
                    dz[sT[i]].Qtkkpb = data[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sT[i];
                    ndz.Qtkkpb = data[i];
                    dz.Add(sT[i], ndz);
                }
            }
            LegendEnviroment.loadRlylPeiBi(Convert.ToInt32(gaolu.Text), "钛球", sjBegin.Value, sjEnd.Value, out sT, out data);

            //钛球配比

            for (int i = 0; i < sT.Length; i++)
            {

                if (dz.ContainsKey(sT[i]))
                {
                    dz[sT[i]].Tqpb = data[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sT[i];
                    ndz.Tqpb = data[i];
                    dz.Add(sT[i], ndz);
                }
            }

            LegendEnviroment.loadRONGJIPeiBi(Convert.ToInt32(gaolu.Text), sjBegin.Value, sjEnd.Value, out sT, out data);

            //熔剂配比

            for (int i = 0; i < sT.Length; i++)
            {

                if (dz.ContainsKey(sT[i]))
                {
                    dz[sT[i]].Rjpb = data[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sT[i];
                    ndz.Rjpb = data[i];
                    dz.Add(sT[i], ndz);
                }
            }
            LegendEnviroment.loadSJHSL(Convert.ToInt32(gaolu.Text), sjBegin.Value, sjEnd.Value, out sT, out data);

            //实际回收率

            for (int i = 0; i < sT.Length; i++)
            {

                if (dz.ContainsKey(sT[i]))
                {
                    dz[sT[i]].Sjhsl = data[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sT[i];
                    ndz.Sjhsl = data[i];
                    dz.Add(sT[i], ndz);
                }
            }
            LegendEnviroment.loadLlhsl(Convert.ToInt32(gaolu.Text), sjBegin.Value, sjEnd.Value, out sT, out data);
            //理论回收率
            for (int i = 0; i < sT.Length; i++)
            {

                if (dz.ContainsKey(sT[i]))
                {
                    dz[sT[i]].Zhpw = data[i];
                    dz[sT[i]].Llhsl = data[i] / 0.946 * 0.997;
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sT[i];
                    ndz.Zhpw = data[i];
                    ndz.Llhsl = data[i] / 0.946 * 0.997;
                    dz.Add(sT[i], ndz);
                }
            }
            LegendEnviroment.loadLDYL(Convert.ToInt32(gaolu.Text), sjBegin.Value, sjEnd.Value, out sT, out data);

            //平均顶压

            for (int i = 0; i < sT.Length; i++)
            {

                if (dz.ContainsKey(sT[i]))
                {
                    dz[sT[i]].Pjdy = data[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sT[i];
                    ndz.Pjdy = data[i];
                    dz.Add(sT[i], ndz);
                }
            }
            LegendEnviroment.loadFYL(Convert.ToInt32(gaolu.Text), sjBegin.Value, sjEnd.Value, out sT, out data);

            //富氧率

            for (int i = 0; i < sT.Length; i++)
            {

                if (dz.ContainsKey(sT[i]))
                {
                    dz[sT[i]].Fyl = data[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sT[i];
                    ndz.Fyl = data[i];
                    dz.Add(sT[i], ndz);
                }
            }
            // weiwancheng
            LegendEnviroment.loadFESI(Convert.ToInt32(gaolu.Text), sjBegin.Value, sjEnd.Value, out sT, out data);

            //生铁含硅

            for (int i = 0; i < sT.Length; i++)
            {

                if (dz.ContainsKey(sT[i]))
                {
                    dz[sT[i]].Fesi = data[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sT[i];
                    ndz.Fesi = data[i];
                    dz.Add(sT[i], ndz);
                }
            }



            //P负荷


            double[] sjkS;
            // 判断是否是5高炉的烧结矿
            if (Convert.ToInt32(gaolu.Text) != 5)
            {
                LegendEnviroment.loadSJKPFH(Convert.ToInt32(gaolu.Text), sjBegin.Value, sjEnd.Value, out sT, out sjkS);

            }
            else
            {
                LegendEnviroment.loadSJKPFH5(Convert.ToInt32(gaolu.Text), sjBegin.Value, sjEnd.Value, out sT, out sjkS);
            }
            for (int i = 0; i < sT.Length; i++)
            {

                if (dz.ContainsKey(sT[i]))
                {
                    dz[sT[i]].Sfh = sjkS[i];     //目前只有烧结的成分
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sT[i];
                    ndz.Sfh = sjkS[i];
                    dz.Add(sT[i], ndz);
                }
            }
            double[] slqS;
            // 判断是否是5高炉的球团
            if (Convert.ToInt32(gaolu.Text) != 5)
            {
                LegendEnviroment.loadQTKPFH(Convert.ToInt32(gaolu.Text), sjBegin.Value, sjEnd.Value, out sT, out slqS);

            }
            else
            {
                LegendEnviroment.loadQTKPFH5(Convert.ToInt32(gaolu.Text), sjBegin.Value, sjEnd.Value, out sT, out slqS);
            }
            // 没有澳矿的时候 只算烧结矿和竖炉
            for (int i = 0; i < sT.Length; i++)
            {

                if (dz.ContainsKey(sT[i]))
                {
                    if (slqS.Length < i)
                    {

                        dz[sT[i]].Sfh = sjkS[i] + slqS[i];     //目前只有烧结矿和竖炉球的成分
                    }
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sT[i];
                    ndz.Sfh = slqS[i] + sjkS[i];
                    dz.Add(sT[i], ndz);
                }
            }
            double[] aokS;
            // 判断澳矿中含P

            LegendEnviroment.loadAOKPFH(Convert.ToInt32(gaolu.Text), sjBegin.Value, sjEnd.Value, out sT, out aokS);



            for (int i = 0; i < sT.Length; i++)
            {

                if (dz.ContainsKey(sT[i]))
                {
                    if (sjkS.Length < i)
                    {
                        dz[sT[i]].Sfh = slqS[i] + sjkS[i] + aokS[i];     //目前只有烧结矿和竖炉球的成分
                    }
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sT[i];
                    if (sjkS.Length < i)
                    {
                        ndz.Sfh = slqS[i] + sjkS[i] + aokS[i];
                    }
                    dz.Add(sT[i], ndz);
                }
            }

            LegendEnviroment.loadlzjd(Convert.ToInt32(gaolu.Text), sjBegin.Value, sjEnd.Value, out sT, out data);

            //炉渣碱度

            for (int i = 0; i < sT.Length; i++)
            {

                if (dz.ContainsKey(sT[i]))
                {
                    dz[sT[i]].Lzjd = data[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sT[i];
                    ndz.Lzjd = data[i];
                    dz.Add(sT[i], ndz);
                }
            }
            LegendEnviroment.loadRcl(Convert.ToInt32(gaolu.Text), sjBegin.Value, sjEnd.Value, out sT, out data);

            //日产量

            for (int i = 0; i < sT.Length; i++)
            {

                if (dz.ContainsKey(sT[i]))
                {
                    dz[sT[i]].Rcl = data[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sT[i];
                    ndz.Rcl = data[i];
                    dz.Add(sT[i], ndz);
                }
            }



            LegendEnviroment.loadJB(Convert.ToInt32(gaolu.Text), sjBegin.Value, sjEnd.Value, out sT, out data);

            //焦比

            for (int i = 0; i < sT.Length; i++)
            {
                if (sT.Length > 0)
                {

                    if (dz.ContainsKey(sT[i]))
                    {
                        dz[sT[i]].Jb = data[i];
                    }
                    else
                    {
                        DataZong ndz = new DataZong();
                        ndz.Rq = sT[i];
                        ndz.Jb = data[i];
                        dz.Add(sT[i], ndz);
                    }
                }
            }

            LegendEnviroment.loadRLB(Convert.ToInt32(gaolu.Text), sjBegin.Value, sjEnd.Value, out sT, out data);

            //燃料比

            for (int i = 0; i < sT.Length; i++)
            {
                if (sT.Length > 0)
                {

                    if (dz.ContainsKey(sT[i]))
                    {
                        dz[sT[i]].Rlb = data[i];
                    }
                    else
                    {
                        DataZong ndz = new DataZong();
                        ndz.Rq = sT[i];
                        ndz.Rlb = data[i];
                        dz.Add(sT[i], ndz);
                    }
                }
            }
            LegendEnviroment.loadMB(Convert.ToInt32(gaolu.Text), sjBegin.Value, sjEnd.Value, out sT, out data);

            //煤比

            for (int i = 0; i < sT.Length; i++)
            {
                if (sT.Length > 0)
                {

                    if (dz.ContainsKey(sT[i]))
                    {
                        dz[sT[i]].Mb = data[i];
                    }
                    else
                    {
                        DataZong ndz = new DataZong();
                        ndz.Rq = sT[i];
                        ndz.Mb = data[i];
                        dz.Add(sT[i], ndz);
                    }
                }
            }
            LegendEnviroment.loadRFWD(Convert.ToInt32(gaolu.Text), sjBegin.Value, sjEnd.Value, out sT, out data);

            //热风温度

            for (int i = 0; i < sT.Length; i++)
            {

                if (dz.ContainsKey(sT[i]))
                {
                    dz[sT[i]].Rfwd = data[i];
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sT[i];
                    ndz.Rfwd = data[i];
                    dz.Add(sT[i], ndz);
                }
            }
          
             LegendEnviroment.loadHTYL(Convert.ToInt32(gaolu.Text), sjBegin.Value, sjEnd.Value, out sT, out data); 
            //含铁原料
            for (int i = 0; i < sT.Length; i++)
            {

                if (dz.ContainsKey(sT[i]))
                {
                    dz[sT[i]].Htyl = data[i];
                
                }
                else
                {
                    DataZong ndz = new DataZong();
                    ndz.Rq = sT[i];
                    ndz.Htyl = data[i];
                    dz.Add(sT[i], ndz);
                }
            
            
            }


                this.c1FlexGrid_tiezha.Rows.Count = dz.Count + 1;

                        for (int i = 0; i < dz.Count; i++)
                        {
                            if (sT.Length > 0 && i < sT.Length)
                                                    {
                                                        DataZong ndz = new DataZong();
                                                        ndz.Num = i + 1;              //序号
                                                        this.c1FlexGrid_tiezha[i + 1, 0] = i + 1;
                                                        ndz.Rq = dz[sT[i]].Rq;   //日期
                                                        this.c1FlexGrid_tiezha[i + 1, 1] = dz[sT[i]].Rq;
                                                        ndz.Zichanjthf = dz[sT[i]].Zichanjthf;   //自产焦炭灰分
                                                        this.c1FlexGrid_tiezha[i + 1, 2] = dz[sT[i]].Zichanjthf;
                                                        ndz.Zichanjtsf = dz[sT[i]].Zichanjtsf;   //自产焦炭水分
                                                        this.c1FlexGrid_tiezha[i + 1, 3] = dz[sT[i]].Zichanjtsf;
                                                        ndz.Waigoujthf = dz[sT[i]].Waigoujthf;   //外购焦炭灰分
                                                        this.c1FlexGrid_tiezha[i + 1, 4] = dz[sT[i]].Waigoujthf;
                                                        ndz.Waigoujtsf = dz[sT[i]].Waigoujtsf;   //外购焦炭水分
                                                        this.c1FlexGrid_tiezha[i + 1, 5] = dz[sT[i]].Waigoujtsf;
                                                        ndz.Pcmfhf = dz[sT[i]].Pcmfhf;   //喷吹煤粉灰分
                                                        this.c1FlexGrid_tiezha[i + 1, 6] = dz[sT[i]].Pcmfhf;
                                                        ndz.Pcmfsf = dz[sT[i]].Pcmfsf;   //喷吹煤粉水分
                                                        this.c1FlexGrid_tiezha[i + 1, 7] = dz[sT[i]].Pcmfsf;
                                                        ndz.Pcmfhff = dz[sT[i]].Pcmfhff;   //喷吹煤粉挥发分
                                                        this.c1FlexGrid_tiezha[i + 1, 8] = dz[sT[i]].Pcmfhff;
                                                        ndz.Sjkpw = dz[sT[i]].Sjkpw;   //烧结矿品味
                                                        this.c1FlexGrid_tiezha[i + 1, 9] = dz[sT[i]].Sjkpw;
                                                        ndz.Qtkpw = dz[sT[i]].Qtkpw;   //球团矿品味
                                                        this.c1FlexGrid_tiezha[i + 1, 10] = dz[sT[i]].Qtkpw;
                                                        ndz.Akpw = dz[sT[i]].Akpw;   //澳矿品味
                                                        this.c1FlexGrid_tiezha[i + 1, 11] = dz[sT[i]].Akpw;
                                                        ndz.Qtkkpw = dz[sT[i]].Qtkkpw;   //其它块矿品味
                                                        this.c1FlexGrid_tiezha[i + 1, 12] = dz[sT[i]].Qtkkpw;
                                                        ndz.Tqpw = dz[sT[i]].Tqpw;   //钛球品味
                                                        this.c1FlexGrid_tiezha[i + 1, 13] = dz[sT[i]].Tqpw;
                                                        ndz.Sjkpb = dz[sT[i]].Sjkpb;   //烧结矿配比
                                                        this.c1FlexGrid_tiezha[i + 1, 14] = dz[sT[i]].Sjkpb;
                                                        ndz.Qhkpb = dz[sT[i]].Qhkpb;   //球团矿配比
                                                        this.c1FlexGrid_tiezha[i + 1, 15] = dz[sT[i]].Qhkpb;
                                                        ndz.Akpb = dz[sT[i]].Akpb;   //澳矿配比
                                                        this.c1FlexGrid_tiezha[i + 1, 16] = dz[sT[i]].Akpb;
                                                        ndz.Qtkkpb = dz[sT[i]].Qtkkpb;   //其它块矿配比
                                                        this.c1FlexGrid_tiezha[i + 1, 17] = dz[sT[i]].Qtkkpb;
                                                        ndz.Tqpb = dz[sT[i]].Tqpb;   //钛球配比
                                                        this.c1FlexGrid_tiezha[i + 1, 18] = dz[sT[i]].Tqpb;
                                                        ndz.Zhpw = dz[sT[i]].Zhpw;   //综合品味
                                                        this.c1FlexGrid_tiezha[i + 1, 19] = dz[sT[i]].Zhpw.ToString("#0.00");
                                                        ndz.Rjpb = dz[sT[i]].Rjpb;   //熔剂配比
                                                        this.c1FlexGrid_tiezha[i + 1, 20] = dz[sT[i]].Rjpb;
                                                        ndz.Rfwd = dz[sT[i]].Rfwd;   //热风温度
                                                        this.c1FlexGrid_tiezha[i + 1, 21] = dz[sT[i]].Rfwd;
                                                        ndz.Pjdy = dz[sT[i]].Pjdy;   //平均顶压
                                                        this.c1FlexGrid_tiezha[i + 1, 22] = dz[sT[i]].Pjdy;
                                                        ndz.Fyl = dz[sT[i]].Fyl;   //富氧率
                                                        this.c1FlexGrid_tiezha[i + 1, 23] = dz[sT[i]].Fyl;
                                                        ndz.Fesi = dz[sT[i]].Fesi;   //生铁含si
                                                        this.c1FlexGrid_tiezha[i + 1, 24] = dz[sT[i]].Fesi;
                                                        ndz.Sfh = dz[sT[i]].Sfh;   //S负荷
                                                        this.c1FlexGrid_tiezha[i + 1, 25] = dz[sT[i]].Sfh;
                                                        ndz.Lzjd = dz[sT[i]].Lzjd;   //炉渣碱度
                                                        this.c1FlexGrid_tiezha[i + 1, 26] = dz[sT[i]].Lzjd;
                                                        ndz.Rcl = dz[sT[i]].Rcl;   //日产品
                                                        this.c1FlexGrid_tiezha[i + 1, 27] = dz[sT[i]].Rcl;
                                                        ndz.Llhsl = dz[sT[i]].Llhsl;   //理论回收率
                                                        this.c1FlexGrid_tiezha[i + 1, 28] = ndz.Llhsl.ToString("#0.00");
                                                        ndz.Sjhsl = dz[sT[i]].Sjhsl;   //实际回收率
                                                        this.c1FlexGrid_tiezha[i + 1, 29] = dz[sT[i]].Sjhsl.ToString("#0.00");
                                                        ndz.Jb = dz[sT[i]].Jb;   //焦比
                                                        this.c1FlexGrid_tiezha[i + 1, 30] = dz[sT[i]].Jb;
                                                        ndz.Rlb = dz[sT[i]].Rlb;   //燃料比
                                                        this.c1FlexGrid_tiezha[i + 1, 31] = dz[sT[i]].Rlb;
                                                        ndz.Mb = dz[sT[i]].Mb;   //煤比
                                                        this.c1FlexGrid_tiezha[i + 1, 32] = dz[sT[i]].Mb;
                                                        ndz.Lltl = ndz.Llhsl * dz[sT[i]].Htyl / 100;   //理论铁量
                                                        this.c1FlexGrid_tiezha[i + 1, 33] = ndz.Lltl;
                                                    }                           
                        }
                     

                     this.c1FlexGrid_tiezha.Subtotal(AggregateEnum.Clear);
                        for (int c = 2; c < c1FlexGrid_tiezha.Cols.Count; c++)
                        {
                            c1FlexGrid_tiezha.Subtotal(AggregateEnum.Average,0, -1, c, "平均值");

                        }

                        c1FlexGrid_tiezha.Refresh();
                        
        }

        private void 保存SToolStripButton_Click(object sender, EventArgs e)
        {
            if (this.saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                c1FlexGrid_tiezha.SaveExcel(this.saveFileDialog1.FileName,FileFlags.IncludeFixedCells);
            }
        }

        private void saveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }
      
    }
}