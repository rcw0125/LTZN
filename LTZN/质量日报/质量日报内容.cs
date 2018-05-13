using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;
using System.Data.OracleClient;

namespace LTZN.质量日报
{
   public class 质量日报内容 : BindingList<质量日报内容项>
    {

        public DateTime riqi;
         public 质量日报内容(DateTime riqi)
        {
            this.riqi=riqi;
        }
         public 质量日报内容项 Get质量日报内容项(string MC)
           {
               质量日报内容项 rx = null;
               foreach (质量日报内容项 x in this)
               {
                   if (x.MC == MC )
                       rx = x;
               }
               return rx;
           }
    }
}
