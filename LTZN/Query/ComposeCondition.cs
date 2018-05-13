using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.Query
{
    public class ComposeCondition : Condition
    {
        public ComposeCondition(Condition d1,LogicOpType logicOp,Condition d2)
        {
            this.Condition1 = d1;
            this.LogicOp = logicOp;
            this.Condition2 = d2;
        }

        public Condition Condition2
        {
            get;
            set;
        }

        public Condition Condition1
        {
            get;
            set;
        }

        /// <summary>
        /// 逻辑关系
        /// </summary>
        public LogicOpType LogicOp
        {
            get;
            set;
        }
    
        public override string ToSql()
        {
            if( LogicOp==LogicOpType.并且)
            {
                return "(" + Condition1.ToSql() + " AND " + Condition2.ToSql() + ")";
            }
            else
            {
                return "(" + Condition1.ToSql() + " OR " + Condition2.ToSql() + ")";
            }
        }
    }
}
