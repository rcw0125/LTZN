using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Data.OracleClient;
using System.Security.Principal;
using System.Threading;

namespace LTZN
{
    public delegate void UserChangeEventHandler(string quan);

    public partial class 炼铁智能主窗体 : Form
    {

        public 炼铁智能主窗体()
        {
            InitializeComponent();
            LtznUserManager.instance.RegisterHandler(this, instance_UserChanged);
            UDPListener.instance.MsgEvent += new MsgEventDelegate(instance_MsgEvent);
            UDPListener.instance.MsgDealEvent += new MsgDealEventDelegate(instance_MsgDealEvent);
        }

        void instance_MsgDealEvent(int count)
        {
            this.Invoke(new MsgDealEventDelegate(MsgDeal), count);
        }

        void instance_MsgEvent(MsgContent msg)
        {
            this.Invoke(new MsgEventDelegate(MsgShow), msg);
        }
        void MsgShow(MsgContent msg)
        {
            MsgContentForm frm = new MsgContentForm(msg);
            frm.Show();
        }

        void MsgDeal(int count)
        {
            if (count > 0)
            {
                tsMsg.Text = "有" + count.ToString() + "条消息未处理";
                tsMsg.BackColor = Color.Red;
            }
            else
            {
                tsMsg.Text = "";
                tsMsg.BackColor = SystemColors.Control;
            }
        }

        void instance_UserChanged(LtznUser ltznUser)
        {
            IPrincipal p = LtznUserManager.instance.CurrentUser;
            if (p != null)
            {
                this.toolStripStatusLabel1.Text = p.Identity.Name;

              if (p.IsInRole("2#大烧"))
               {
                   生产.Visible = false;
                   报表.Visible = false;
                   数据分析ToolStripMenuItem.Visible = false;
                   高炉燃料比综合分析系统.Visible = false;
                   查询.Visible = false;
               }
               if (p.IsInRole("1#大烧"))
               {
                    生产.Visible = false;
                    报表.Visible = false;
                    数据分析ToolStripMenuItem.Visible = false;
                    高炉燃料比综合分析系统.Visible = false;
                    查询.Visible = false;
                   
             }
                if (p.IsInRole("统计"))
                {
                    报表参数ToolStripMenuItem.Enabled = true;
                    

                }
                if (p.IsInRole("管理员"))
                {
                    参数设定ToolStripMenuItem.Enabled = true;
                    用户管理ToolStripMenuItem.Enabled = true;
                    报表参数ToolStripMenuItem.Enabled = true;
                    出铁计划ToolStripMenuItem.Enabled = true;
                    出铁方案ToolStripMenuItem.Enabled = true;
                    铁次管理ToolStripMenuItem.Enabled = true;
                    料线探齿ToolStripMenuItem.Enabled = true;
                }
                if (p.IsInRole("1高炉") || p.IsInRole("3高炉") || p.IsInRole("5高炉") || p.IsInRole("6高炉"))
                {
                    参数设定ToolStripMenuItem.Enabled = true;
                }
                if (p.IsInRole("计算模型"))
                {
                    btnModelManager.Enabled = true;
                }
                else
                {
                    btnModelManager.Enabled = false;
                }

                登陆.Text = "退出";
                修改密码.Visible = true;
            }
            else
            {
                参数设定ToolStripMenuItem.Enabled = false;
                用户管理ToolStripMenuItem.Enabled = false;
                报表参数ToolStripMenuItem.Enabled = false;
                出铁计划ToolStripMenuItem.Enabled = false;
                出铁方案ToolStripMenuItem.Enabled = false;
                铁次管理ToolStripMenuItem.Enabled = false;
                料线探齿ToolStripMenuItem.Enabled = false;
                btnModelManager.Enabled = false;
                this.toolStripStatusLabel1.Text = "";
                登陆.Text = "登陆";
                修改密码.Visible = false;
            }
            UDPListener.instance.resetUnDealMsg();
        }

        private void 炼铁智能主窗体_Load(object sender, EventArgs e)
        {
            UDPListener.instance.Start();
        }

        //private IDockContent FindDocument(string text)
        //{
        //    if (dockPanel.DocumentStyle == DocumentStyle.SystemMdi)
        //    {
        //        foreach (Form form in MdiChildren)
        //            if (form.Text == text)
        //                return form as IDockContent;

        //        return null;
        //    }
        //    else
        //    {
        //        foreach (IDockContent content in dockPanel.Contents)
        //            if (content.TabId == text)
        //                return content;

        //        return null;
        //    }

        //}

        //public void ShowDockForm(Type type, string tabText, DockState dockState, params object[] args)
        //{
        //    IDockContent dc = FindDocument(tabText);
        //    if (dc != null)
        //    {
        //        dc.DockHandler.Activate();
        //        return;
        //    }
        //    DockContent frm = Activator.CreateInstance(type, args) as DockContent;
        //    if (dockPanel.DocumentStyle == DocumentStyle.SystemMdi)
        //    {
        //        frm.MdiParent = this;
        //        frm.Show();
        //    }
        //    else
        //    {
        //        frm.Show(dockPanel, dockState);
        //    }
        //}

        //public void TransferDockMessage(DockContentMessage dockContentMessage)
        //{
        //    foreach (IDealDockContentMessage ddMsg in dockPanel.Contents)
        //    {
        //        ddMsg.DealDockContentMessage(dockContentMessage);
        //    }
        //}

        private Form FindForm(string text)
        {
            foreach (Form form in MdiChildren)
                if (form.Text == text)
                    return form;
            return null;
        }

        public void ShowForm(Type type, string tabText, params object[] args)
        {
            this.Cursor = Cursors.WaitCursor;
            Form dc = FindForm(tabText);
            if (dc != null)
            {
                dc.WindowState = FormWindowState.Maximized;
                dc.Activate();
                this.Cursor = Cursors.Default;
                return;
            }

            Form frm = Activator.CreateInstance(type, args) as Form;
            frm.MdiParent = this;
            frm.WindowState = FormWindowState.Maximized;
            frm.Show();
            this.Cursor = Cursors.Default;
        }

        private void 高炉日消耗_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(XiaohaoGaolu.XiaohaoGaolu), "高炉日消耗录入");
        }

        private void 全厂日消耗_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(全厂日消耗.全厂日消耗窗体), "全厂日消耗窗体");
        }

        private void 技术日报_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(技术日报.RiReportForm), "技术日报模型");
        }
        private void 技术月报_Click(object sender, EventArgs e)
        {
           // ShowForm(typeof(技术月报.Form1), "技术月报");
            ShowForm(typeof(技术月报.MonthReportForm),"炼铁技术月报");
        }

        private void 原料消耗_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(原料消耗.原料消耗), "原料消耗");
        }

        private void 登陆ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            if (登陆.Text == "登陆")
            {
                登陆窗体 a = new 登陆窗体();
                if (a.ShowDialog() == DialogResult.OK)
                {
                    if (!LtznUserManager.instance.Login(a.yonghu, a.mima))
                        MessageBox.Show("登陆失败，请确认用户名密码！");
                }
            }
            else
            {
                LtznUserManager.instance.Quit();
            }
            this.Cursor = Cursors.Default;
        }

        private void 调度_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(调度.调度窗体), "调度窗体");
        }

        private void 查询_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(查询.数据查询窗口), "数据查询窗口");
        }

        private void 生产_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(生产日报.生产日报), "生产日报");
        }

        private void 修改密码_Click(object sender, EventArgs e)
        {
            if (LtznUserManager.instance.CurrentUser != null)
            {
                密码更改窗体 a = new 密码更改窗体();
                a.yonghu = LtznUserManager.instance.CurrentUser.Identity.Name;
                a.ShowDialog();
            }
        }

        private void 月报调整_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(技术月报.月报调整), "月报调整");
        }

        private void 技术年报ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(技术年报.年报浏览窗体), "年报浏览窗体");
        }

        private void 参数设定ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            数据分析.参数设定 fm = new 数据分析.参数设定();
            fm.WindowState = FormWindowState.Normal;
            fm.ShowDialog(this);
            this.Cursor = Cursors.Default;
        }

        private void 炉温ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(数据分析.炉温), "炉温");
        }

        private void 剩余热ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(数据分析.剩余热), "剩余热");
        }

        private void 碱度趋势ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(数据分析.碱度趋势), "碱度趋势");
        }

        private void 用户管理ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            用户管理 fm = new 用户管理();
            fm.WindowState = FormWindowState.Normal;
            fm.ShowDialog(this);
            this.Cursor = Cursors.Default;
        }

      

        private void 产量分析ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(数据分析.产量), "产量");
        }

        private void 发送信息ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (LtznUserManager.instance.CurrentUser != null)
            {
                MsgSendForm frm = new MsgSendForm();
                frm.Show(this);
            }
        }

        private void 炼铁智能主窗体_FormClosed(object sender, FormClosedEventArgs e)
        {

            UDPListener.instance.Stop();
        }

        private void unDealMsgForm_show(object sender, EventArgs e)
        {
            if (LtznUserManager.instance.CurrentUser != null)
            {
                MsgUnDealForm frm = new MsgUnDealForm();
                frm.ShowDialog(this);
            }
        }

        private void 报表_Click(object sender, EventArgs e)
        {

        }

        private void 报表参数ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ParaForm frm = new ParaForm();
            frm.ShowDialog();
        }

        private void 原料分析ToolStripMenuItem_Click(object sender, EventArgs e)
        {
           ShowForm(typeof(数据分析.MaterialAnalysisForm), "原料分析");
        }

        private void 计算模型ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(CalModelForm), "计算模型");
        }

        private void 技术日报模型ToolStripMenuItem_Click(object sender, EventArgs e)
        {
          
        }

        private void 出铁数据ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(数据分析.出铁数据), "出铁数据");
        }

        private void 关于ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox frm = new AboutBox();
            frm.ShowDialog();
        }

        private void 燃料质量分析_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(高炉燃料比综合分析.燃料质量分析), "燃料质量分析");
        }

        private void 原料品位分析ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(高炉燃料比综合分析.原料品味分析), "原料品位分析");
        }

        private void 入炉原料结构ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(高炉燃料比综合分析.入炉原料结构), "入炉原料结构");
        }

        private void 风温富氧分析ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(高炉燃料比综合分析.风温富氧率分析), "风温富氧率分析");
        }

        private void 理论回收率分析ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(高炉燃料比综合分析.回收率分析), "回收率分析");
        }

        private void 燃料消耗ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(高炉燃料比综合分析.燃料消耗分析), "燃料消耗分析");
        }

        private void 热风压力分析ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(高炉燃料比综合分析.热风炉顶压力分析), "热风炉顶压力分析");
        }

        private void 数据汇总分析ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(高炉燃料比综合分析.数据汇总), "数据汇总");
        }

        private void 澳矿参数设定ToolStripMenuItem_Click(object sender, EventArgs e)
        {
          //  ShowForm(typeof(高炉燃料比综合分析.澳矿成分录入), "澳矿成分录入");

            高炉燃料比综合分析.澳矿成分录入 frm = new 高炉燃料比综合分析.澳矿成分录入();
            frm.ShowDialog();
        }

        private void mN趋势ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void cr趋势ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void si趋势ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(数据分析.炉温), "炉温");
        }

        private void s趋势ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(数据分析.S趋势), "S趋势");
        }

        private void ti趋势ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(数据分析.TiO2), "TiO2趋势");
        }

        private void 趋势参数ToolStripMenuItem_Click(object sender, EventArgs e)
        {
           
            数据分析.趋势参数 fm = new 数据分析.趋势参数();
            fm.WindowState = FormWindowState.Normal;
            fm.ShowDialog(this);
            this.Cursor = Cursors.Default;
        }

        private void p趋势ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(数据分析.P趋势), "P趋势");
        }

        private void ni趋势ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(数据分析.Ni趋势), "Ni趋势");
        }

        private void cr趋势ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(数据分析.Cr趋势), "Cr趋势");
        }

        private void mn趋势ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(数据分析.Mn趋势), "Mn趋势");
        }

        private void s趋势ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(数据分析.烧结矿S趋势), "烧结矿S趋势");
        }

        private void p趋势ToolStripMenuItem_Click(object sender, EventArgs e)
        {

            ShowForm(typeof(数据分析.烧结矿P趋势), "烧结矿P趋势");
        }

        private void si趋势ToolStripMenuItem2_Click(object sender, EventArgs e)
        {

            ShowForm(typeof(数据分析.烧结矿Si趋势), "烧结矿Si趋势");
        }

        private void ti趋势ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(数据分析.烧结矿Ti趋势), "烧结矿Ti趋势");
        }

        private void 原料设定ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            数据分析.原料分类 fm = new 数据分析.原料分类();
            fm.WindowState = FormWindowState.Normal;
            fm.ShowDialog(this);
            this.Cursor = Cursors.Default;
        }

        private void 更新内容ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            更新内容 gx = new 更新内容();
            gx.Show();
        }

        private void 出铁方案ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            出铁方案 form = new 出铁方案();
            form.Show();
        }

        private void 出铁计划ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            出铁计划 form1 = new 出铁计划();
            form1.Show();
        }

        private void 铁次管理ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void 质量日报ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(质量日报.质量日报Form), "技术日报模型");
        }

        private void 物理热ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(数据分析.物理热趋势), "物理热趋势");
        }

        private void s趋势ToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(数据分析.newS趋势), "newS趋势");
        }

        private void 碱度ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ShowForm(typeof(数据分析.new碱度趋势), "new碱度趋势");
        }

        private void 料线探齿ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            料线探齿设置 form = new 料线探齿设置();
            form.Show();
        }

  
      

       

    }
}
