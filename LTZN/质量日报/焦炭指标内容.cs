using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;
using System.Data.OracleClient;
namespace LTZN.质量日报
{
    public class 焦炭指标内容 : BindingList<焦炭指标内容项>
    {

           public DateTime riqi;
        public 焦炭指标内容(DateTime riqi)
        {
            this.riqi=riqi;
        }
        public 焦炭指标内容()
        {
            this.riqi = DateTime.Today;
        }












    }
}
