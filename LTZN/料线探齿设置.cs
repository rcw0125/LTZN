using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace LTZN
{
    public partial class 料线探齿设置 : Form
    {
        public 料线探齿设置()
        {
            InitializeComponent();
        }

        private void 料线探齿设置_Load(object sender, EventArgs e)
        {
            // TODO: 这行代码将数据加载到表“lxbz.LXBIAOZHI”中。您可以根据需要移动或移除它。
            this.lXBIAOZHITableAdapter.Fill(this.lxbz.LXBIAOZHI);

        }

        private void 保存_Click(object sender, EventArgs e)
        {
            bool keyibaocun = true;
            for (int i=0; i < this.lxbz.LXBIAOZHI.Count;i++ )
            {
                if (this.lxbz.LXBIAOZHI[i].LIAOXIAN1 == this.lxbz.LXBIAOZHI[i].LIAOXIAN2)
                {

                    MessageBox.Show("只能使用一条料线");
                    keyibaocun = false;

                }

                if ((this.lxbz.LXBIAOZHI[i].LIAOXIAN1 != 0 && this.lxbz.LXBIAOZHI[i].LIAOXIAN1 != 1) || (this.lxbz.LXBIAOZHI[i].LIAOXIAN2 != 0 && this.lxbz.LXBIAOZHI[i].LIAOXIAN2 != 1))

                {


                    MessageBox.Show("不符合填写标准");
                    keyibaocun = false;
                
                }
            }

            if (keyibaocun == true)
            {
                this.lXBIAOZHITableAdapter.Update(this.lxbz.LXBIAOZHI);
                MessageBox.Show("保存完成");
            }
        }
    }
}
