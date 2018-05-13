using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;
using System.Data.OracleClient;

namespace LTZN.质量日报
{
    public class 套筒窑内容 : BindingList<套筒窑内容项>
    {
          public DateTime riqi;
          public 套筒窑内容(DateTime riqi)
        {
            this.riqi=riqi;
        }
          public 套筒窑内容项 Get套筒窑内容项(string MC)
           {
               套筒窑内容项 rx = null;
               foreach (套筒窑内容项 x in this)
               {
                   if (x.MC == MC )
                       rx = x;
               }
               return rx;
           }
    }
}
