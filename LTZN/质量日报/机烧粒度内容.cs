using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;
using System.Data.OracleClient;
namespace LTZN.质量日报
{
   public class 机烧粒度内容 : BindingList<机烧粒度内容项>
    {

        
        public DateTime riqi;
        public 机烧粒度内容(DateTime riqi)
        {
            this.riqi=riqi;
        }
        public 机烧粒度内容()
        {
            this.riqi = DateTime.Today;
        }
        public 机烧粒度内容项 Get质量日报内容项(string MC, DateTime SJ)
           {
               机烧粒度内容项 rx = null;
               foreach (机烧粒度内容项 x in this)
               {
                   if (x.MC == MC && x.SJ == SJ)
                       rx = x;
               }
               return rx;
           }







    }
}
