using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;
using System.Data.OracleClient;
namespace LTZN.技术日报
{
    public class 技术日报内容 : BindingList<技术日报内容项>
    {

        public DateTime riqi;

        public 技术日报内容(DateTime riqi)
        {
            this.riqi=riqi;
        }
        public 技术日报内容()
        {
            this.riqi = DateTime.Today;
        }
        public 技术日报内容项 Get技术日报内容项(string 单位, string 项目)
        {
            技术日报内容项 rx = null;
            foreach (技术日报内容项 x in this)
            {
                if (x.P01单位 == 单位 && x.P02项目 == 项目)
                    rx = x;
            }
            return rx;
        }
    }
}
