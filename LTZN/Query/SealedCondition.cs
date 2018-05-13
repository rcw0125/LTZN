using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.Query
{
    public abstract class SealedCondition : Condition
    {

        public string FieldName
        {
            get;
            set;
        }

        public RelationOpType RalationOp
        {
            get;
            set;
        }
    
        public override string ToSql()
        {
            throw new NotImplementedException();
        }
    }
}
