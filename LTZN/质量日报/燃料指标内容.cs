using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;
using System.Data.OracleClient;
namespace LTZN.质量日报
{
    public class 燃料指标内容 : BindingList<燃料指标内容项>
    {

         public DateTime riqi;
        public 燃料指标内容(DateTime riqi)
        {
            this.riqi=riqi;
        }
        public 燃料指标内容()
        {
            this.riqi = DateTime.Today;
        }
        public 燃料指标内容项 Get燃料指标内容项(string MC, DateTime SJ)
           {
               燃料指标内容项 rx = null;
               foreach (燃料指标内容项 x in this)
               {
                   if (x.MC == MC && x.SJ == SJ)
                       rx = x;
               }
               return rx;
           }
    }
}
