using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.Query
{
    public abstract class Condition
    {
        public abstract string ToSql();
    }
}
