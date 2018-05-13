using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.CalFramework
{
    public class CalUtility
    {
        StringBuilder StrB;
        
        private int iCurr = 0;
        private int iCount = 0;
     
        /// <summary>
        /// 构造方法
        /// </summary>
        public CalUtility(string calStr)
        {
            StrB = new StringBuilder(calStr.Trim());
            iCount= StrB.Length;
          //  iCount = System.Text.Encoding.Default.GetByteCount(calStr.Trim());
        }

        /// <summary>
        /// 取段，自动分析数值或计算符
        /// </summary>
        /// <returns></returns>
        public FormaItem getItem()
        {
            FormaItem result = new FormaItem();
            //结束了
            if (iCurr == iCount)
            {
                result.Name = "";
                result.Typ = FormaItemType.EndForma;
                return result;
            }

            char ChTmp = StrB[iCurr];
            if (IsOper(ChTmp))
            {
                iCurr++;
                result.Name = ChTmp.ToString();
                result.Typ = FormaItemType.Operator;
                return result;
            }

            result.Typ = FormaItemType.Number;
            StringBuilder strTmp = new StringBuilder();
            while (true)
            {

               // if (char.IsLetterOrDigit(ChTmp) || ChTmp == '.')
                if(!char.IsControl(ChTmp) && !char.IsWhiteSpace(ChTmp))
                {
                    strTmp.Append(ChTmp);

                    if (!IsNum(ChTmp))
                        result.Typ = FormaItemType.Tag;
                }

                iCurr++;

                if (iCurr == iCount)
                    break;
                ChTmp = StrB[iCurr];
                if (IsOper(ChTmp))
                {
                    if (strTmp.Length == 0)
                    {
                        strTmp.Append(ChTmp);
                        result.Typ = FormaItemType.Operator;
                        iCurr++;
                    }
                    break;
                }
            }


            result.Name = strTmp.ToString();
            if (result.Name == "")
                result.Typ = FormaItemType.EndForma;
            return result;
        }
        /// <summary>
        /// 是否是数字
        /// </summary>
        /// <param name="c">内容</param>
        /// <returns></returns>
        public bool IsNum(char c)
        {
            if ((c >= '0' && c <= '9') || c == '.')
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// 是否是数字
        /// </summary>
        /// <param name="c">内容</param>
        /// <returns></returns>
        public bool IsNum(string c)
        {
            if (c.Equals(""))
                return false;
            if ((c[0] >= '0' && c[0] <= '9') || c[0] == '.')
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// 是否是运算符
        /// </summary>
        /// <param name="c">内容</param>
        /// <returns></returns>
        public bool IsOper(char c)
        {
            if (c == '+' || c == '-' || c == '*' || c == '/' || c == '(' || c == ')')
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        /// <summary>
        /// 是否是运算符
        /// </summary>
        /// <param name="c">内容</param>
        /// <returns></returns>
        public bool IsOper(string c)
        {
            if (c == "+" || c == "-" || c == "*" || c == "/" || c == "(" || c == ")" )
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        
    }

    public struct FormaItem
    {
        public string Name;
        public FormaItemType Typ;
    }

    public enum FormaItemType
    {
        Operator,
        Number,
        Tag,
        EndForma
    }
}
