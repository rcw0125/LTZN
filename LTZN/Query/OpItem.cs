using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.Query
{
    public class OpItem
    {
        public OpItem(OpState state)
        {
            this.OpState = state;
            this.OpValue = null;
        }

        public OpItem(OpState state,object opval)
        {
            this.OpState = state;
            this.OpValue = opval;
        }
    
        public OpState OpState
        {
            get;
            set;
        }

        public object OpValue
        {
            get;
            set;
        }
    }
}
