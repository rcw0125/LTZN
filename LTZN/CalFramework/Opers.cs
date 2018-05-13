using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.CalFramework
{
    /// <summary>
    /// Opers 的摘要说明。
    /// 各类计算符的接口实现，加减乘除
    /// </summary>
    public class OperAdd : IOper
    {
        public OperAdd()
        {
            //
            // TODO: 在此处添加构造函数逻辑
            //
        }
        #region IOper 成员
        public object Oper(object o1, object o2)
        {
            Decimal d1 = Decimal.Parse(o1.ToString());
            Decimal d2 = Decimal.Parse(o2.ToString());
            return d1 + d2;
        }
        #endregion
    }
    public class OperDec : IOper
    {
        public OperDec()
        {
            //
            // TODO: 在此处添加构造函数逻辑
            //
        }
        #region IOper 成员
        public object Oper(object o1, object o2)
        {
            Decimal d1 = Decimal.Parse(o1.ToString());
            Decimal d2 = Decimal.Parse(o2.ToString());
            return d1 - d2;
        }
        #endregion
    }
    public class OperRide : IOper
    {
        public OperRide()
        {
            //
            // TODO: 在此处添加构造函数逻辑
            //
        }
        #region IOper 成员
        public object Oper(object o1, object o2)
        {
            Decimal d1 = Decimal.Parse(o1.ToString());
            Decimal d2 = Decimal.Parse(o2.ToString());
            return d1 * d2;
        }
        #endregion
    }
    public class OperDiv : IOper
    {
        public OperDiv()
        {
            //
            // TODO: 在此处添加构造函数逻辑
            //
        }
        #region IOper 成员
        public object Oper(object o1, object o2)
        {
            Decimal d1 = Decimal.Parse(o1.ToString());
            Decimal d2 = Decimal.Parse(o2.ToString());
            if (d2 == 0)
                return 0;
            else
                return d1 / d2;
        }
        #endregion
    }
}
