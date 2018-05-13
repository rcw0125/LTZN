using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.Query
{
    public enum OpState
    {
        开始,
        左括号,
        逻辑关系,
        条件,
        右括号,
        结束,
    }
}
