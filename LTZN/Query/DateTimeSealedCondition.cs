using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.Query
{
    public class DateTimeSealedCondition : SealedCondition
    {
        public DateTime FieldValue
        {
            get;
            set;
        }

        public override string ToSql()
        {
            string result = "";
            switch (this.RalationOp)
            {
                case RelationOpType.大于:
                    result = this.FieldName + " > TO_DATE('" + this.FieldValue.ToString() + "','YYYY-MM-DD HH24:MI:SS')";
                    break;
                case RelationOpType.大于等于:
                    result = this.FieldName + " >= TO_DATE('" + this.FieldValue + "','YYYY-MM-DD HH24:MI:SS')";
                    break;
                case RelationOpType.小于:
                    result = this.FieldName + "< TO_DATE('" + this.FieldValue + "','YYYY-MM-DD HH24:MI:SS')";
                    break;
                case RelationOpType.小于等于:
                    result = this.FieldName + " <= TO_DATE('" + this.FieldValue + "','YYYY-MM-DD HH24:MI:SS')";
                    break;
                case RelationOpType.等于:
                    result = this.FieldName + " = TO_DATE('" + this.FieldValue + "','YYYY-MM-DD HH24:MI:SS')";
                    break;
                case RelationOpType.不等于:
                    result = this.FieldName + " <> TO_DATE('" + this.FieldValue + "','YYYY-MM-DD HH24:MI:SS')";
                    break;
                case RelationOpType.是空:
                    result = this.FieldName + " IS NULL";
                    break;
                case RelationOpType.非空:
                    result = this.FieldName + " IS NOT NULL"; ;
                    break;
                default:
                    break;
            }
            return result;
        }
    }

}