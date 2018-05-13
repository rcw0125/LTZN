using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.CalFramework
{
    public class AggregateCal : CalTagDecorator
    {

        private List<CalTag> _subItems = new List<CalTag>();
       
        public List<CalTag> SubItems
        {
            get {
                return _subItems;
            }
        }

        public override void Cal()
        {
            if (!CalFinished)
            {
                double total = 0;
                foreach (var item in SubItems)
                {
                    total += item.TagValue;
                }
                this.TagValue = total;
                CalFinished = true;
            }
        }

        public AggregateCal(CalTag tag)
            : base(tag)
        {

        }
    }
}
