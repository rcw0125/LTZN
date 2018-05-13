using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.Query
{
    public class EntityField
    {
        public string FieldName { get; set; }

        public string FieldDispaly { get; set; }

        public string FieldType { get; set; }

        public string ColList { get; set; }

        public SealedCondition Create(string relationOpStr, object val)
        {
            
            if (FieldType == "string")
            {
                StringSealedCondition condition = new StringSealedCondition();
                condition.FieldName = this.FieldName;
                condition.FieldValue = (string)val;
                condition.RalationOp = (RelationOpType)Enum.Parse(typeof(RelationOpType), relationOpStr);
                return condition;
            }
            if (FieldType == "DateTime")
            {
                DateTimeSealedCondition condition = new DateTimeSealedCondition();
                condition.FieldName = this.FieldName;
                if (val is string)
                {
                    condition.FieldValue = DateTime.Parse((string)val);//(DateTime)val;
                }
                else if (val is DateTime)
                {
                    condition.FieldValue = (DateTime)val;
                }
                condition.RalationOp = (RelationOpType)Enum.Parse(typeof(RelationOpType), relationOpStr);
                return condition;
            }
            if (FieldType == "double")
            {
                NumericSealedCondition condition = new NumericSealedCondition();
                condition.FieldName = this.FieldName;
                if (val is string)
                {
                    condition.FieldValue = double.Parse((string)val);// (double)val;
                }
                else if (val is double)
                {
                    condition.FieldValue =(double)val;
                }
                condition.RalationOp = (RelationOpType)Enum.Parse(typeof(RelationOpType), relationOpStr);
                return condition;
            } 
             if (FieldType == "bool" || FieldType == "int")
            {
                IntSealedCondition condition = new IntSealedCondition();
                condition.FieldName = this.FieldName;
                if (val is int)
                {
                    condition.FieldValue = (int)val;// (double)val;
                } 
                else if (val is string)
                {
                    string strVal=val as string;
                    if (strVal.ToLower() == "true" || strVal == "是")
                    {
                        condition.FieldValue = 1;
                    }
                    else if (strVal.ToLower() == "false" || strVal == "否")
                    {
                        condition.FieldValue = 0;
                    }
                    else
                    {
                        condition.FieldValue = int.Parse((string)val);// (double)val;
                    }
                }
                else if (val is bool)
                {
                    if ((bool)val)
                    {
                        condition.FieldValue = 1;
                    }
                    else
                    {
                        condition.FieldValue = 0;
                    }
                }
                else if (val is double)
                {
                    condition.FieldValue = Convert.ToInt32(val);
                }
                
                condition.RalationOp = (RelationOpType)Enum.Parse(typeof(RelationOpType), relationOpStr);
                return condition;
            }
            return null;
        }
    }
}
