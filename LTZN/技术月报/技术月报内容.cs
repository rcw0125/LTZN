using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;
using System.Windows.Forms;

namespace LTZN.技术月报
{
    public class 技术月报内容 : BindingList<技术月报内容项>
    {

        public int nian;
        public int yue;

        public 技术月报内容(int nian, int yue)
        {
            this.nian = nian;
            this.yue = yue;
        }

        public 技术月报内容项 Get技术月报内容项(string 单位, string 项目)
        {
            技术月报内容项 rx = null;
            foreach (技术月报内容项 x in this)
            {
                if (x.P01单位 == 单位 && x.P02项目 == 项目)
                    rx = x;
            }
            return rx;
        }
    }
}
