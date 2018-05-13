using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.CalFramework2
{
    /// <summary>
    /// OperFactory 的摘要说明。
    /// 计算符接口工厂
    /// </summary>
    public class OperFactory
    {
        public OperFactory()
        {
        }

        public IOper CreateOper(string Oper)
        {
            if (Oper.Equals("+"))
            {
                IOper p = new OperAdd();
                return p;
            }
            if (Oper.Equals("-"))
            {
                IOper p = new OperDec();
                return p;
            }
            if (Oper.Equals("*"))
            {
                IOper p = new OperRide();
                return p;
            }
            if (Oper.Equals("/"))
            {
                IOper p = new OperDiv();
                return p;
            }
            return null;
        }

        /// <summary>
        /// 比较str1和str2两个运算符的优先级，ture表示str1高于str2，false表示str1低于str2
        /// </summary>
        /// <param name="str1">计算符1</param>
        /// <param name="str2">计算符2</param>
        /// <returns></returns>
        public bool MoreThen(string str1, string str2)
        {
            return getPriority(str1) >= getPriority(str2);
        }

        /// <summary>
        /// 取得计算符号的优先级
        /// </summary>
        /// <param name="str">计算符</param>
        /// <returns></returns>
        public int getPriority(string str)
        {
            if (str.Equals(""))
            {
                return -1;
            }
            if (str.Equals(")"))
            {
                return 0;
            }
            if (str.Equals("+") || str.Equals("-"))
            {
                return 1;
            }
            if (str.Equals("*") || str.Equals("/"))
            {
                return 2;
            }
            if (str.Equals("("))
            {
                return 3;
            }
            return 0;
        }
    }
}
