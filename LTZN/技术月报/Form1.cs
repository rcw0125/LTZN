using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace LTZN.技术月报
{
    public partial class Form1 : Form
    {
        private bool canUpdate = false;

        public Form1()
        {
            InitializeComponent();
            LtznUserManager.instance.RegisterHandler(this, Form1_UserChange);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.TextBox_nian.Text = DateTime.Today.Year.ToString();
            this.TextBox_yue.Text = DateTime.Today.Month.ToString();
            getData(Convert.ToInt32(this.TextBox_nian.Text), Convert.ToInt32(this.TextBox_yue.Text));
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (canUpdate)
            {
                this.技术月报内容BindingSource.CurrencyManager.EndCurrentEdit();
                this.技术月报TableAdapter1.Update(this.技术月报表1.技术月报);
            }
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            if (canUpdate)
            {
                this.技术月报内容BindingSource.CurrencyManager.EndCurrentEdit();
                this.技术月报TableAdapter1.Update(this.技术月报表1.技术月报);
            }
            技术月报窗体 f = new 技术月报窗体();
            f.WindowState = FormWindowState.Maximized;
            f.显示报表(Convert.ToInt32(TextBox_nian.Text), Convert.ToInt32(TextBox_yue.Text));
            this.Cursor = Cursors.Default;
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            try
            {
                getData(Convert.ToInt32(this.TextBox_nian.Text),Convert.ToInt32(this.TextBox_yue.Text));

            }
            catch
            {
            }
            this.Cursor = Cursors.Default;
        }

        private void getData(int nian, int yue)
        {
            if (canUpdate)
            {
                this.技术月报内容BindingSource.CurrencyManager.EndCurrentEdit();
                this.技术月报TableAdapter1.Update(this.技术月报表1.技术月报);
            }
            this.技术月报TableAdapter1.Fill(this.技术月报表1.技术月报, nian, yue);

            if (this.技术月报表1.技术月报.Count != 12)
            {
                if (this.技术月报表1.技术月报.Count > 0)
                {
                    foreach (技术月报表.技术月报Row r in this.技术月报表1.技术月报)
                    {
                        r.Delete();
                    }
                    this.技术月报TableAdapter1.Update(this.技术月报表1.技术月报);

                }
                this.技术月报表1.技术月报.Add技术月报Row(nian, yue, "1#", "本月");
                this.技术月报表1.技术月报.Add技术月报Row(nian, yue, "1#", "累计");
                this.技术月报表1.技术月报.Add技术月报Row(nian, yue, "2#", "本月");
                this.技术月报表1.技术月报.Add技术月报Row(nian, yue, "2#", "累计");
                this.技术月报表1.技术月报.Add技术月报Row(nian, yue, "3#", "本月");
                this.技术月报表1.技术月报.Add技术月报Row(nian, yue, "3#", "累计");
                this.技术月报表1.技术月报.Add技术月报Row(nian, yue, "4#", "本月");
                this.技术月报表1.技术月报.Add技术月报Row(nian, yue, "4#", "累计");
                this.技术月报表1.技术月报.Add技术月报Row(nian, yue, "5#", "本月");
                this.技术月报表1.技术月报.Add技术月报Row(nian, yue, "5#", "累计");
                this.技术月报表1.技术月报.Add技术月报Row(nian, yue, "Q", "本月");
                this.技术月报表1.技术月报.Add技术月报Row(nian, yue, "Q", "累计");
            }

            this.技术月报内容BindingSource.DataSource = this.技术月报表1.技术月报;
        }

        private void C1NumericEdit_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                (sender as Control).Parent.SelectNextControl((sender as Control),true, true,false, true);
        }

        //从日报提取
        private void button2_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            DateTime rq = new DateTime(Convert.ToInt32(TextBox_nian.Text), Convert.ToInt32(TextBox_yue.Text), 1);
            月报参数 yxh = new 月报参数();
            yxh.getDataBy(rq);

            for (int i = 0; i < 6; i++)
            {

                技术月报表.技术月报Row r = this.技术月报表1.技术月报[i * 2];
                r.P03全铁产量 =(decimal) yxh.全铁产量[i];
                r.P04合格铁 =(decimal) yxh.合格铁[i];
                r.P05制钢铁 =(decimal) yxh.制钢铁[i];
                r.P06铸造铁 =(decimal) yxh.铸造铁[i];
                r.P07折合产量 =(decimal) yxh.折算产量[i];
                r.P08合格率 =(decimal) yxh.合格率[i];
                r.P09一级品率 =(decimal) yxh.一级品率[i];
                r.P10送炼钢优质铁水率 =(decimal) yxh.送炼钢优质铁水率[i];
                r.P11高炉利用系数 =(decimal) yxh.高炉利用系数[i];
                r.P12入炉焦炭冶炼强度 =(decimal) yxh.入炉焦炭冶炼强度[i];
                r.P13综合焦炭冶炼强度 =(decimal) yxh.综合焦炭冶炼强度[i];
                r.P14折算综合焦比 =(decimal) yxh.折算综合焦比[i];
                r.P15折合入炉焦比 =(decimal) yxh.折合入炉焦比[i];
                r.P16入炉矿品位 =(decimal) yxh.入炉综合品位[i];
                r.P17熟料比 =(decimal) yxh.熟料比[i];
                r.P18入炉焦炭负荷 =(decimal) yxh.入炉焦炭负荷[i];
                r.P19综合焦炭负荷 =(decimal) yxh.综合焦炭负荷[i];
                r.P20休风时间 =(decimal) yxh.休风小时[i];
                r.P21休风率 =(decimal) yxh.休风率[i];
                r.P22慢风时间 =(decimal) yxh.慢风[i];
                r.P23慢风率 =(decimal) yxh.慢风率[i];
                r.P24煤气成分CO2 =(decimal) yxh.CO2[i];
                r.P25煤气成分计算1 =(decimal) yxh.CO2率[i];
                r.P27深料线作业率 =(decimal) yxh.深料线率[i];
                r.P28回收率 =(decimal) yxh.回收率[i];
                r.P29冷风流量 =(decimal) yxh.冷风流量[i];
                r.P30平均风温 =(decimal) yxh.平均风温[i];
                r.P32热风压力 =(decimal) yxh.热风压力[i];
                r.P33炉顶压力 =(decimal) yxh.炉顶压力[i];
                r.P34炉顶温度 =(decimal) yxh.炉顶温度[i];
                r.P35富氧率 =(decimal) yxh.富氧率[i];
                r.P36高压率 =(decimal) yxh.高压率[i];
                r.P37悬料次数 =(decimal) yxh.悬料次数[i];
                r.P38坐料次数 =(decimal) yxh.坐料次数[i];
                r.P39崩料次数 =(decimal) yxh.崩料次数[i];
                r.P40风口大套损坏数 =(decimal) yxh.风口损坏大[i];
                r.P41风口中套损坏数 =(decimal) yxh.风口损坏中[i];
                r.P42风口小套损坏数 =(decimal) yxh.风口损坏小[i];
                r.P43渣口大套损坏数 =(decimal) yxh.渣口损坏大[i];
                r.P44渣口中套损坏数 =(decimal) yxh.渣口损坏中[i];
                r.P45渣口小套损坏数 =(decimal) yxh.渣口损坏小[i];
                r.P46渣铁比 =(decimal) yxh.渣铁比[i];
                r.P47灰铁比 =(decimal) yxh.灰铁比[i];
                r.P48原料总耗 =(decimal) yxh.原料总耗[i];
                r.P49原料单耗 =(decimal) yxh.原料单耗[i];
                r.P50机烧消耗 =(decimal) yxh.机烧[i];
                r.P51竖炉球消耗 =(decimal) yxh.竖球[i];
                r.P53其它熟料消耗 =(decimal) yxh.其它熟料[i];
                r.P54本溪矿消耗 =(decimal) yxh.本溪矿[i];
                r.P55其它生料消耗 =(decimal) yxh.其它生料[i];
                r.P56废铁总耗 =(decimal) yxh.废铁[i];
                r.P57废铁单耗 =(decimal) yxh.废铁单耗[i];
                r.P58金属收得率 =(decimal) yxh.金属收得率[i];
                r.P59综合焦炭总耗 =(decimal) yxh.综合焦炭[i];
                r.P60七号称 =(decimal) yxh.七号称焦炭[i];
                r.P61干毛焦 =(decimal) yxh.干毛焦[i];
                r.P62干焦粉 =(decimal) yxh.干焦粉[i];
                r.P63入炉焦炭总耗 =(decimal) yxh.入炉焦炭[i];
                r.P64入炉焦炭单耗 =(decimal) yxh.入炉焦炭单耗[i];
                r.P65煤粉总耗 =(decimal) yxh.煤粉[i];
                r.P66焦丁总耗 =(decimal) yxh.焦丁[i];
                r.P67焦丁单耗 =(decimal) yxh.焦丁单耗[i];
                r.P68燃料比 =(decimal) yxh.燃料比[i];
                r.P69铁成分SI =(decimal) yxh.铁Si[i];
                r.P70铁成分MN =(decimal) yxh.铁Mn[i];
                r.P71铁成分P =(decimal) yxh.铁P[i];
                r.P72铁成分S =(decimal) yxh.铁S[i];
                r.P73渣成分CAO =(decimal) yxh.渣CaO[i];
                r.P74渣成分SIO2 =(decimal) yxh.渣SiO2[i];
                r.P75渣成分AL2O3 =(decimal) yxh.渣Al2O3[i];
                r.P76渣成分MGO =(decimal) yxh.渣MgO[i];
                r.P77渣成分FEO =(decimal) yxh.渣FeO[i];
                r.P78渣成分S =(decimal) yxh.渣S[i];
                r.P79含SI偏差CAO =(decimal) yxh.碱度[i];
                r.P80含SI偏差制铁量 =(decimal) yxh.制钢铁Si偏差[i];
                r.P81含SI偏差铸造铁 =(decimal) yxh.铸造铁Si偏差[i];
                r.综合焦炭单耗 =(decimal) yxh.综合焦炭单耗[i];
                r.煤粉单耗 =(decimal) yxh.煤粉单耗[i];
                r.一级铁 =(decimal) yxh.一级铁[i];
                r.优质铁 =(decimal) yxh.送炼钢优质铁[i];
                r.瓦斯灰 =(decimal) yxh.瓦斯灰[i];
                r.收入含铁 = (decimal)yxh.收入含铁[i];
                r.支出含铁 = (decimal)yxh.支出含铁[i];
                r.富氧量 = (decimal)yxh.富氧流量[i];
                r.高压操作时间 = (decimal)yxh.高压操作时间[i];
                r.有效工作时间 = (decimal)yxh.有效工作时间[i];
                r.理论渣量 = (decimal)yxh.理论渣量[i];
                r.深料线 = (decimal)yxh.深料线[i];
                r.全部料线 = (decimal)yxh.全部料线[i];
                r.高炉有效容积 = (decimal)yxh.高炉有效炉容[i];
                r.高炉实际容积 = (decimal)yxh.高炉实际炉容[i];
                r.入炉矿含铁 = (decimal)yxh.入炉矿含铁[i];
                r.P82工艺称焦比 = (decimal)yxh.工艺称焦比[i];
            }
            if (rq.Month == 1)
            {
                for (int i = 0; i < 6; i++)
                {

                    技术月报表.技术月报Row r = this.技术月报表1.技术月报[i * 2+1];
                    r.P03全铁产量 = (decimal)yxh.全铁产量[i];
                    r.P04合格铁 = (decimal)yxh.合格铁[i];
                    r.P05制钢铁 = (decimal)yxh.制钢铁[i];
                    r.P06铸造铁 = (decimal)yxh.铸造铁[i];
                    r.P07折合产量 = (decimal)yxh.折算产量[i];
                    r.P08合格率 = (decimal)yxh.合格率[i];
                    r.P09一级品率 = (decimal)yxh.一级品率[i];
                    r.P10送炼钢优质铁水率 = (decimal)yxh.送炼钢优质铁水率[i];
                    r.P11高炉利用系数 = (decimal)yxh.高炉利用系数[i];
                    r.P12入炉焦炭冶炼强度 = (decimal)yxh.入炉焦炭冶炼强度[i];
                    r.P13综合焦炭冶炼强度 = (decimal)yxh.综合焦炭冶炼强度[i];
                    r.P14折算综合焦比 = (decimal)yxh.折算综合焦比[i];
                    r.P15折合入炉焦比 = (decimal)yxh.折合入炉焦比[i];
                    r.P16入炉矿品位 = (decimal)yxh.入炉综合品位[i];
                    r.P17熟料比 = (decimal)yxh.熟料比[i];
                    r.P18入炉焦炭负荷 = (decimal)yxh.入炉焦炭负荷[i];
                    r.P19综合焦炭负荷 = (decimal)yxh.综合焦炭负荷[i];
                    r.P20休风时间 = (decimal)yxh.休风小时[i];
                    r.P21休风率 = (decimal)yxh.休风率[i];
                    r.P22慢风时间 = (decimal)yxh.慢风[i];
                    r.P23慢风率 = (decimal)yxh.慢风率[i];
                    r.P24煤气成分CO2 = (decimal)yxh.CO2[i];
                    r.P25煤气成分计算1 = (decimal)yxh.CO2率[i];
                    r.P27深料线作业率 = (decimal)yxh.深料线率[i];
                    r.P28回收率 = (decimal)yxh.回收率[i];
                    r.P29冷风流量 = (decimal)yxh.冷风流量[i];
                    r.P30平均风温 = (decimal)yxh.平均风温[i];
                    r.P32热风压力 = (decimal)yxh.热风压力[i];
                    r.P33炉顶压力 = (decimal)yxh.炉顶压力[i];
                    r.P34炉顶温度 = (decimal)yxh.炉顶温度[i];
                    r.P35富氧率 = (decimal)yxh.富氧率[i];
                    r.P36高压率 = (decimal)yxh.高压率[i];
                    r.P37悬料次数 = (decimal)yxh.悬料次数[i];
                    r.P38坐料次数 = (decimal)yxh.坐料次数[i];
                    r.P39崩料次数 = (decimal)yxh.崩料次数[i];
                    r.P40风口大套损坏数 = (decimal)yxh.风口损坏大[i];
                    r.P41风口中套损坏数 = (decimal)yxh.风口损坏中[i];
                    r.P42风口小套损坏数 = (decimal)yxh.风口损坏小[i];
                    r.P43渣口大套损坏数 = (decimal)yxh.渣口损坏大[i];
                    r.P44渣口中套损坏数 = (decimal)yxh.渣口损坏中[i];
                    r.P45渣口小套损坏数 = (decimal)yxh.渣口损坏小[i];
                    r.P46渣铁比 = (decimal)yxh.渣铁比[i];
                    r.P47灰铁比 = (decimal)yxh.灰铁比[i];
                    r.P48原料总耗 = (decimal)yxh.原料总耗[i];
                    r.P49原料单耗 = (decimal)yxh.原料单耗[i];
                    r.P50机烧消耗 = (decimal)yxh.机烧[i];
                    r.P51竖炉球消耗 = (decimal)yxh.竖球[i];
                    r.P53其它熟料消耗 = (decimal)yxh.其它熟料[i];
                    r.P54本溪矿消耗 = (decimal)yxh.本溪矿[i];
                    r.P55其它生料消耗 = (decimal)yxh.其它生料[i];
                    r.P56废铁总耗 = (decimal)yxh.废铁[i];
                    r.P57废铁单耗 = (decimal)yxh.废铁单耗[i];
                    r.P58金属收得率 = (decimal)yxh.金属收得率[i];
                    r.P59综合焦炭总耗 = (decimal)yxh.综合焦炭[i];
                    r.P60七号称 = (decimal)yxh.七号称焦炭[i];
                    r.P61干毛焦 = (decimal)yxh.干毛焦[i];
                    r.P62干焦粉 = (decimal)yxh.干焦粉[i];
                    r.P63入炉焦炭总耗 = (decimal)yxh.入炉焦炭[i];
                    r.P64入炉焦炭单耗 = (decimal)yxh.入炉焦炭单耗[i];
                    r.P65煤粉总耗 = (decimal)yxh.煤粉[i];
                    r.P66焦丁总耗 = (decimal)yxh.焦丁[i];
                    r.P67焦丁单耗 = (decimal)yxh.焦丁单耗[i];
                    r.P68燃料比 = (decimal)yxh.燃料比[i];
                    r.P69铁成分SI = (decimal)yxh.铁Si[i];
                    r.P70铁成分MN = (decimal)yxh.铁Mn[i];
                    r.P71铁成分P = (decimal)yxh.铁P[i];
                    r.P72铁成分S = (decimal)yxh.铁S[i];
                    r.P73渣成分CAO = (decimal)yxh.渣CaO[i];
                    r.P74渣成分SIO2 = (decimal)yxh.渣SiO2[i];
                    r.P75渣成分AL2O3 = (decimal)yxh.渣Al2O3[i];
                    r.P76渣成分MGO = (decimal)yxh.渣MgO[i];
                    r.P77渣成分FEO = (decimal)yxh.渣FeO[i];
                    r.P78渣成分S = (decimal)yxh.渣S[i];
                    r.P79含SI偏差CAO = (decimal)yxh.碱度[i];
                    r.P80含SI偏差制铁量 = (decimal)yxh.制钢铁Si偏差[i];
                    r.P81含SI偏差铸造铁 = (decimal)yxh.铸造铁Si偏差[i];
                    r.综合焦炭单耗 = (decimal)yxh.综合焦炭单耗[i];
                    r.煤粉单耗 = (decimal)yxh.煤粉单耗[i];
                    r.一级铁 = (decimal)yxh.一级铁[i];
                    r.优质铁 = (decimal)yxh.送炼钢优质铁[i];
                    r.瓦斯灰 = (decimal)yxh.瓦斯灰[i];
                    r.收入含铁 = (decimal)yxh.收入含铁[i];
                    r.支出含铁 = (decimal)yxh.支出含铁[i];
                    r.富氧量 = (decimal)yxh.富氧流量[i];
                    r.高压操作时间 = (decimal)yxh.高压操作时间[i];
                    r.有效工作时间 = (decimal)yxh.有效工作时间[i];
                    r.理论渣量 = (decimal)yxh.理论渣量[i];
                    r.深料线 = (decimal)yxh.深料线[i];
                    r.全部料线 = (decimal)yxh.全部料线[i];
                    r.高炉有效容积 = (decimal)yxh.高炉有效炉容[i];
                    r.高炉实际容积 = (decimal)yxh.高炉实际炉容[i];
                    r.入炉矿含铁 = (decimal)yxh.入炉矿含铁[i];
                    r.P82工艺称焦比 = (decimal)yxh.工艺称焦比[i];
                }
            }
            else
            {
                月报参数 ybcs = new 月报参数();
                上月累计 sylj = new 上月累计();
                if (rq.Month > 1)
                    sylj.getDataBy(rq.Year, rq.Month - 1);
                for (int i = 0; i < 6; i++)
                {
                    int a = 0;
                    int b = 0;
                    if (sylj.制钢铁Si偏差[i] > 0)
                        a = 1;
                    else
                        a = 0;
                    if (yxh.制钢铁Si偏差[i] > 0)
                        b = 1;
                    else
                        b = 0;
                    if (sylj.制钢铁[i] * a + yxh.制钢铁[i] * b > 0)
                        sylj.制钢铁Si偏差[i] = (sylj.制钢铁Si偏差[i] * sylj.制钢铁[i] + yxh.制钢铁Si偏差[i] * yxh.制钢铁[i]) / (sylj.制钢铁[i] * a + yxh.制钢铁[i] * b);
                    else
                        sylj.制钢铁Si偏差[i] = 0;
                    sylj.制钢铁Si偏差[i] = double.Parse(sylj.制钢铁Si偏差[i].ToString("N2"));

                    if (sylj.铸造铁Si偏差[i] > 0)
                        a = 1;
                    else
                        a = 0;
                    if (yxh.铸造铁Si偏差[i] > 0)
                        b = 1;
                    else
                        b = 0;
                    if (sylj.铸造铁[i] * a + yxh.铸造铁[i] * b > 0)
                        sylj.铸造铁Si偏差[i] = (sylj.铸造铁Si偏差[i] * sylj.铸造铁[i] + yxh.铸造铁Si偏差[i] * yxh.铸造铁[i]) / (sylj.铸造铁[i] * a + yxh.铸造铁[i] * b);
                    else
                        sylj.铸造铁Si偏差[i] = 0;

                    sylj.铸造铁Si偏差[i] = double.Parse(sylj.铸造铁Si偏差[i].ToString("N2"));
                    sylj.全铁产量[i] += yxh.全铁产量[i];
                    sylj.合格铁[i] += yxh.合格铁[i];
                    sylj.制钢铁[i] += yxh.制钢铁[i];
                    sylj.一级铁[i] += yxh.一级铁[i];
                    sylj.送炼钢优质铁[i] += yxh.送炼钢优质铁[i];
                    sylj.折算产量[i] += yxh.折算产量[i];

                    sylj.煤粉[i] += yxh.煤粉[i];
                    sylj.焦丁[i] += yxh.焦丁[i];
                    sylj.干毛焦[i] += yxh.干毛焦[i];
                    sylj.入炉焦炭[i] += yxh.入炉焦炭[i];
                    sylj.干焦粉[i] += yxh.干焦粉[i];
                    sylj.七号称焦炭[i] += yxh.七号称焦炭[i];

                    sylj.原料总耗[i] += yxh.原料总耗[i];
                    sylj.机烧[i] += yxh.机烧[i];
                    sylj.竖球[i] += yxh.竖球[i];
                    sylj.其它熟料[i] += yxh.其它熟料[i];
                    sylj.本溪矿[i] += yxh.本溪矿[i];
                    sylj.其它生料[i] += yxh.其它生料[i];

                    sylj.废铁[i] += yxh.废铁[i];
                    sylj.瓦斯灰[i] += yxh.瓦斯灰[i];

                    sylj.悬料次数[i] += yxh.悬料次数[i];
                    sylj.坐料次数[i] += yxh.坐料次数[i];
                    sylj.崩料次数[i] += yxh.崩料次数[i];

                    sylj.风口损坏大[i] += yxh.风口损坏大[i];
                    sylj.风口损坏中[i] += yxh.风口损坏中[i];
                    sylj.风口损坏小[i] += yxh.风口损坏小[i];
                    sylj.渣口损坏大[i] += yxh.渣口损坏大[i];
                    sylj.渣口损坏中[i] += yxh.渣口损坏中[i];
                    sylj.渣口损坏小[i] += yxh.渣口损坏小[i];

                    sylj.休风小时[i] += yxh.休风小时[i];
                    sylj.慢风[i] += yxh.慢风[i];
                    sylj.冷风流量[i] += yxh.冷风流量[i];
                    sylj.富氧流量[i] += yxh.富氧流量[i];
                    sylj.收入含铁[i] += yxh.收入含铁[i];
                    sylj.支出含铁[i] += yxh.支出含铁[i];
                    sylj.高压操作时间[i] += yxh.高压操作时间[i];
                    sylj.有效工作时间[i] += yxh.有效工作时间[i];
                    sylj.理论渣量[i] += yxh.理论渣量[i];
                    sylj.深料线[i] += yxh.深料线[i];
                    sylj.全部料线[i] += yxh.全部料线[i];
                    sylj.高炉有效炉容[i] += yxh.高炉有效炉容[i];
                    sylj.高炉实际炉容[i] += yxh.高炉实际炉容[i];
                    sylj.铁Si[i] = (sylj.铁Si[i] * sylj.月数 + yxh.铁Si[i]) / (Convert.ToInt32(sylj.铁Si[i] > 0) * sylj.月数 + Convert.ToInt32(yxh.铁Si[i] > 0));
                    sylj.铁Si[i] = format(sylj.铁Si[i],"N2");
                    sylj.铁Mn[i] = (sylj.铁Mn[i] * sylj.月数 + yxh.铁Mn[i]) / (Convert.ToInt32(sylj.铁Mn[i] > 0) * sylj.月数 + Convert.ToInt32(yxh.铁Mn[i] > 0));
                    sylj.铁Mn[i] = format(sylj.铁Mn[i],"N2");
                    sylj.铁P[i] = (sylj.铁P[i] * sylj.月数 + yxh.铁P[i]) / (Convert.ToInt32(sylj.铁P[i] > 0) * sylj.月数 + Convert.ToInt32(yxh.铁P[i] > 0));
                    sylj.铁P[i] = format(sylj.铁P[i],"N3");
                    sylj.铁S[i] = (sylj.铁S[i] * sylj.月数 + yxh.铁S[i]) / (Convert.ToInt32(sylj.铁S[i] > 0) * sylj.月数 + Convert.ToInt32(yxh.铁S[i] > 0));
                    sylj.铁S[i] = format(sylj.铁S[i],"N3");
                    sylj.渣SiO2[i] = (sylj.渣SiO2[i] * sylj.月数 + yxh.渣SiO2[i]) / (Convert.ToInt32(sylj.渣SiO2[i] > 0) * sylj.月数 + Convert.ToInt32(yxh.渣SiO2[i] > 0));
                    sylj.渣SiO2[i] = format(sylj.渣SiO2[i],"N2");
                    sylj.渣CaO[i] = (sylj.渣CaO[i] * sylj.月数 + yxh.渣CaO[i]) / (Convert.ToInt32(sylj.渣CaO[i] > 0) * sylj.月数 + Convert.ToInt32(yxh.渣CaO[i] > 0));
                    sylj.渣CaO[i] = format(sylj.渣CaO[i],"N2");
                    sylj.渣MgO[i] = (sylj.渣MgO[i] * sylj.月数 + yxh.渣MgO[i]) / (Convert.ToInt32(sylj.渣MgO[i] > 0) * sylj.月数 + Convert.ToInt32(yxh.渣MgO[i] > 0));
                    sylj.渣MgO[i] = format(sylj.渣MgO[i],"N2");
                    sylj.渣Al2O3[i] = (sylj.渣Al2O3[i] * sylj.月数 + yxh.渣Al2O3[i]) / (Convert.ToInt32(sylj.渣Al2O3[i] > 0) * sylj.月数 + Convert.ToInt32(yxh.渣Al2O3[i] > 0));
                    sylj.渣Al2O3[i] = format(sylj.渣Al2O3[i],"N2");
                    sylj.渣S[i] = (sylj.渣S[i] * sylj.月数 + yxh.渣S[i]) / (Convert.ToInt32(sylj.渣S[i] > 0) * sylj.月数 + Convert.ToInt32(yxh.渣S[i] > 0));
                    sylj.渣S[i] = format(sylj.渣S[i],"N2");
                    sylj.热风压力[i] = (sylj.热风压力[i] * sylj.月数 + yxh.热风压力[i]) / (Convert.ToInt32(sylj.热风压力[i] > 0) * sylj.月数 + Convert.ToInt32(yxh.热风压力[i] > 0));
                    sylj.热风压力[i] = format(sylj.热风压力[i],"N0");
                    sylj.炉顶压力[i] = (sylj.炉顶压力[i] * sylj.月数 + yxh.炉顶压力[i]) / (Convert.ToInt32(sylj.炉顶压力[i] > 0) * sylj.月数 + Convert.ToInt32(yxh.炉顶压力[i] > 0));
                    sylj.炉顶压力[i] = format(sylj.炉顶压力[i],"N0");
                    sylj.炉顶温度[i] = (sylj.炉顶温度[i] * sylj.月数 + yxh.炉顶温度[i]) / (Convert.ToInt32(sylj.炉顶温度[i] > 0) * sylj.月数 + Convert.ToInt32(yxh.炉顶温度[i] > 0));
                    sylj.炉顶温度[i] = format(sylj.炉顶温度[i],"N0");
                    sylj.CO2[i] = (sylj.CO2[i] * sylj.月数 + yxh.CO2[i]) / (Convert.ToInt32(sylj.CO2[i] > 0) * sylj.月数 + Convert.ToInt32(yxh.CO2[i] > 0));
                    sylj.CO2[i] = format(sylj.CO2[i],"N2");
                    sylj.CO2率[i] = (sylj.CO2率[i] * sylj.月数 + yxh.CO2率[i]) / (Convert.ToInt32(sylj.CO2率[i] > 0) * sylj.月数 + Convert.ToInt32(yxh.CO2率[i]>0));
                    sylj.CO2率[i] = format(sylj.CO2率[i],"N0");
                    sylj.入炉矿含铁[i] += yxh.入炉矿含铁[i];
                    sylj.平均风温[i] = (sylj.平均风温[i] * (sylj.累计天数 - yxh.本月天数) + yxh.平均风温[i] * yxh.本月天数) / (Convert.ToInt32(sylj.平均风温[i] > 0) * (sylj.累计天数 - yxh.本月天数) + Convert.ToInt32(yxh.平均风温[i] > 0) * yxh.本月天数);
                    sylj.平均风温[i] = format(sylj.平均风温[i],"N0");

                }
                sylj.计算();

                for (int i = 0; i < 6; i++)
                {
                    技术月报表.技术月报Row r = this.技术月报表1.技术月报[i * 2 + 1];
                    r.P03全铁产量 = (decimal)sylj.全铁产量[i];
                    r.P04合格铁 = (decimal)sylj.合格铁[i];
                    r.P05制钢铁 = (decimal)sylj.制钢铁[i];
                    r.P06铸造铁 = (decimal)sylj.铸造铁[i];
                    r.P07折合产量 = (decimal)sylj.折算产量[i];
                    r.P08合格率 = (decimal)sylj.合格率[i];
                    r.P09一级品率 = (decimal)sylj.一级品率[i];
                    r.P10送炼钢优质铁水率 = (decimal)sylj.送炼钢优质铁水率[i];
                    r.P11高炉利用系数 = (decimal)sylj.高炉利用系数[i];
                    r.P12入炉焦炭冶炼强度 = (decimal)sylj.入炉焦炭冶炼强度[i];
                    r.P13综合焦炭冶炼强度 = (decimal)sylj.综合焦炭冶炼强度[i];
                    r.P14折算综合焦比 = (decimal)sylj.折算综合焦比[i];
                    r.P15折合入炉焦比 = (decimal)sylj.折合入炉焦比[i];
                    r.P16入炉矿品位 = (decimal)sylj.入炉综合品位[i];
                    r.P17熟料比 = (decimal)sylj.熟料比[i];
                    r.P18入炉焦炭负荷 = (decimal)sylj.入炉焦炭负荷[i];
                    r.P19综合焦炭负荷 = (decimal)sylj.综合焦炭负荷[i];
                    r.P20休风时间 = (decimal)sylj.休风小时[i];
                    r.P21休风率 = (decimal)sylj.休风率[i];
                    r.P22慢风时间 = (decimal)sylj.慢风[i];
                    r.P23慢风率 = (decimal)sylj.慢风率[i];
                    r.P24煤气成分CO2 = (decimal)sylj.CO2[i];
                    r.P25煤气成分计算1 = (decimal)sylj.CO2率[i];
                    r.P27深料线作业率 = (decimal)sylj.深料线率[i];
                    r.P28回收率 = (decimal)sylj.回收率[i];
                    r.P29冷风流量 = (decimal)sylj.冷风流量[i];
                    r.P30平均风温 = (decimal)sylj.平均风温[i];
                    r.P32热风压力 = (decimal)sylj.热风压力[i];
                    r.P33炉顶压力 = (decimal)sylj.炉顶压力[i];
                    r.P34炉顶温度 = (decimal)sylj.炉顶温度[i];
                    r.P35富氧率 = (decimal)sylj.富氧率[i];
                    r.P36高压率 = (decimal)sylj.高压率[i];
                    r.P37悬料次数 = (decimal)sylj.悬料次数[i];
                    r.P38坐料次数 = (decimal)sylj.坐料次数[i];
                    r.P39崩料次数 = (decimal)sylj.崩料次数[i];
                    r.P40风口大套损坏数 = (decimal)sylj.风口损坏大[i];
                    r.P41风口中套损坏数 = (decimal)sylj.风口损坏中[i];
                    r.P42风口小套损坏数 = (decimal)sylj.风口损坏小[i];
                    r.P43渣口大套损坏数 = (decimal)sylj.渣口损坏大[i];
                    r.P44渣口中套损坏数 = (decimal)sylj.渣口损坏中[i];
                    r.P45渣口小套损坏数 = (decimal)sylj.渣口损坏小[i];
                    r.P46渣铁比 = (decimal)sylj.渣铁比[i];
                    r.P47灰铁比 = (decimal)sylj.灰铁比[i];
                    r.P48原料总耗 = (decimal)sylj.原料总耗[i];
                    r.P49原料单耗 = (decimal)sylj.原料单耗[i];
                    r.P50机烧消耗 = (decimal)sylj.机烧[i];
                    r.P51竖炉球消耗 = (decimal)sylj.竖球[i];
                    r.P53其它熟料消耗 = (decimal)sylj.其它熟料[i];
                    r.P54本溪矿消耗 = (decimal)sylj.本溪矿[i];
                    r.P55其它生料消耗 = (decimal)sylj.其它生料[i];
                    r.P56废铁总耗 = (decimal)sylj.废铁[i];
                    r.P57废铁单耗 = (decimal)sylj.废铁单耗[i];
                    r.P58金属收得率 = (decimal)sylj.金属收得率[i];
                    r.P59综合焦炭总耗 = (decimal)sylj.综合焦炭[i];
                    r.P60七号称 = (decimal)sylj.七号称焦炭[i];
                    r.P61干毛焦 = (decimal)sylj.干毛焦[i];
                    r.P62干焦粉 = (decimal)sylj.干焦粉[i];
                    r.P63入炉焦炭总耗 = (decimal)sylj.入炉焦炭[i];
                    r.P64入炉焦炭单耗 = (decimal)sylj.入炉焦炭单耗[i];
                    r.P65煤粉总耗 = (decimal)sylj.煤粉[i];
                    r.P66焦丁总耗 = (decimal)sylj.焦丁[i];
                    r.P67焦丁单耗 = (decimal)sylj.焦丁单耗[i];
                    r.P68燃料比 = (decimal)sylj.燃料比[i];
                    r.P69铁成分SI = (decimal)sylj.铁Si[i];
                    r.P70铁成分MN = (decimal)sylj.铁Mn[i];
                    r.P71铁成分P = (decimal)sylj.铁P[i];
                    r.P72铁成分S = (decimal)sylj.铁S[i];
                    r.P73渣成分CAO = (decimal)sylj.渣CaO[i];
                    r.P74渣成分SIO2 = (decimal)sylj.渣SiO2[i];
                    r.P75渣成分AL2O3 = (decimal)sylj.渣Al2O3[i];
                    r.P76渣成分MGO = (decimal)sylj.渣MgO[i];
                    r.P77渣成分FEO = (decimal)sylj.渣FeO[i];
                    r.P78渣成分S = (decimal)sylj.渣S[i];
                    r.P79含SI偏差CAO = (decimal)sylj.碱度[i];
                    r.P80含SI偏差制铁量 = (decimal)sylj.制钢铁Si偏差[i];
                    r.P81含SI偏差铸造铁 = (decimal)sylj.铸造铁Si偏差[i];
                    r.综合焦炭单耗 = (decimal)sylj.综合焦炭单耗[i];
                    r.煤粉单耗 = (decimal)sylj.煤粉单耗[i];
                    r.一级铁 = (decimal)sylj.一级铁[i];
                    r.优质铁 = (decimal)sylj.送炼钢优质铁[i];
                    r.瓦斯灰 = (decimal)sylj.瓦斯灰[i];
                    r.收入含铁 = (decimal)sylj.收入含铁[i];
                    r.支出含铁 = (decimal)sylj.支出含铁[i];
                    r.富氧量 = (decimal)sylj.富氧流量[i];
                    r.高压操作时间 = (decimal)sylj.高压操作时间[i];
                    r.有效工作时间 = (decimal)sylj.有效工作时间[i];
                    r.理论渣量 = (decimal)sylj.理论渣量[i];
                    r.深料线 = (decimal)sylj.深料线[i];
                    r.全部料线 = (decimal)sylj.全部料线[i];
                    r.高炉有效容积 = (decimal)sylj.高炉有效炉容[i];
                    r.高炉实际容积 = (decimal)sylj.高炉实际炉容[i];
                    r.入炉矿含铁 = (decimal)sylj.入炉矿含铁[i];
                    r.P82工艺称焦比 = (decimal)sylj.工艺称焦比[i];
                }
            }
            this.Cursor = Cursors.Default;
        }

        void Form1_UserChange(LtznUser ltznUser)
        {
            if (ltznUser!=null && ltznUser.IsInRole("统计"))
            {
                canUpdate = true;
                button2.Enabled = true;
            }
            else
            {
                canUpdate = false;
                button2.Enabled = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.技术月报表1.技术月报.计算();
        }
        private double format(double d,string f)
        {
            if (double.IsNaN(d) || double.IsInfinity(d) || d <= 0)
                return 0;
            else
               return  double.Parse(d.ToString(f));
        }

    }
}