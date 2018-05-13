using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.XiaohaoGaolu
{
    internal class XiaohaoManager : List<Xiaohao>
    {
       
        public void Add(DateTime riqi,int gaolu)
        {
            Clear();
            Add(new Xiaohao(riqi, gaolu));
        }
        public void save()
        {
            if (Count > 0)
                base[0].EndEdit();
        }

    }
}
