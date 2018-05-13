using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace LTZN.CalFramework2
{
    public class CalUtility : List<FormulaItem>
    {

        /// <summary>
        /// 构造方法
        /// </summary>
        public CalUtility(string calStr)
        {
            ParseFormula(calStr);
        }

        private void ParseFormula(string formula)
        {
        
            StringBuilder StrB = new StringBuilder(formula.Trim());
            int iCount = StrB.Length;
            //iCount = System.Text.Encoding.Default.GetByteCount(calStr.Trim());

            for (int iCurr = 0; iCurr < iCount; iCurr++)
            {
                FormulaItem result = new FormulaItem();
                result.Typ = FormulaItemType.Number;
                StringBuilder strTmp = new StringBuilder();
                while (true)
                {
                    if (iCurr >= iCount)
                    {
                        break;
                    }

                    char ChTmp = StrB[iCurr];


                    if (strTmp.ToString() == "")
                    {
                        if (IsOper(ChTmp))
                        {
                            strTmp.Append(ChTmp);
                            result.Typ = FormulaItemType.Operator;
                            iCurr++;
                            break;
                        }
                    }

                    if (!char.IsControl(ChTmp) && !char.IsWhiteSpace(ChTmp) && ChTmp != '+' && ChTmp != '*' && ChTmp != '/' && ChTmp != '(' && ChTmp != ')')
                    {
                        if (strTmp.ToString() == "" || ChTmp != '-')
                        {
                            strTmp.Append(ChTmp);

                            if (!IsNum(ChTmp))
                                result.Typ = FormulaItemType.Tag;
                        }
                        else
                        {
                            break;
                        }
                    }
                    else
                    {
                        break;
                    }
                    iCurr++;
                }

                result.Name = strTmp.ToString();
                if (result.Name == "")
                    result.Typ = FormulaItemType.EndForma;
                lastFormaItem = result;
                Add(result);
                if (result.Typ == FormulaItemType.EndForma) break;
            }
        }

        private Nullable<FormulaItem> lastFormaItem = null;

      
        /// <summary>
        /// 是否是数字
        /// </summary>
        /// <param name="c">内容</param>
        /// <returns></returns>
        public bool IsNum(char c)
        {
            if ((c >= '0' && c <= '9') || c == '.' || c == '-')
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
            if ((c[0] >= '0' && c[0] <= '9') || c[0] == '.' || c[0] == '-')
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
            if (c == '+' || c == '*' || c == '/' || c == '(' || c == ')' || c == '=')
            {
                return true;
            }
            else if (c == '-')
            {
                if (!lastFormaItem.HasValue)
                {
                    return false;
                }
                else if (lastFormaItem.Value.Typ == FormulaItemType.Operator && lastFormaItem.Value.Name != ")")
                {
                    return false;
                }
                else
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
            if (c == "+" || c == "*" || c == "/" || c == "(" || c == ")")
            {
                return true;
            }
            else if (c == "-")
            {
                if (!lastFormaItem.HasValue)
                {
                    return false;
                }
                else if (lastFormaItem.Value.Typ == FormulaItemType.Operator && lastFormaItem.Value.Name != ")")
                {
                    return false;
                }
                else
                    return true;
            }
            else
            {
                return false;
            }
        }

        public object DoCal(MemTagModel tm)
        {
            OperFactory of = new OperFactory();
            /// 算术符栈
            ArrayList HList = new ArrayList();
            /// 数值栈
            ArrayList Vlist = new ArrayList();

            foreach (var formaItem in this)
            {
                if (formaItem.Typ == FormulaItemType.Number)
                {
                    //如果是数值，则写入数据栈
                    Vlist.Add(formaItem.Name);
                }
                else if (formaItem.Typ == FormulaItemType.Tag)
                {
                    double? val = tm.GetTagValue(formaItem.Name);
                    Vlist.Add(val == null ? "0" : val.ToString());
                }
                else if (formaItem.Name != "=")
                {
                    string opStr = formaItem.Name;

                    while (HList.Count > 0)
                    {
                        //比较优先级
                        if (of.MoreThen(HList[HList.Count - 1].ToString(), opStr))
                        {
                            if (HList[HList.Count - 1].ToString().Equals("("))
                            {
                                if (opStr.Equals(")")) //符号是否可以对消？
                                {
                                    HList.RemoveAt(HList.Count - 1);
                                    if (HList.Count > 0)
                                    {
                                        opStr = HList[HList.Count - 1].ToString();
                                        HList.RemoveAt(HList.Count - 1);
                                    }
                                }
                                else
                                {
                                    break;
                                }
                            }
                            else
                            {
                                //如果新来操作符优先级低，则计算前面的运算符
                                IOper p = of.CreateOper(HList[HList.Count - 1].ToString());
                                if (p != null)
                                {
                                    Vlist[Vlist.Count - 2] = p.Oper(Vlist[Vlist.Count - 2], Vlist[Vlist.Count - 1]);
                                    HList.RemoveAt(HList.Count - 1);
                                    Vlist.RemoveAt(Vlist.Count - 1);
                                }
                                else
                                {
                                    break;
                                }
                            }
                        }
                        else
                        {
                            break;
                        }
                    }

                    if (!opStr.Equals(""))
                        HList.Add(opStr);
                    else
                        break;
                }
            }
            if (HList.Count > 0)
            {
                IOper p = of.CreateOper(HList[HList.Count - 1].ToString());
                if (p != null)
                {
                    Vlist[Vlist.Count - 2] = p.Oper(Vlist[Vlist.Count - 2], Vlist[Vlist.Count - 1]);
                    HList.RemoveAt(HList.Count - 1);
                    Vlist.RemoveAt(Vlist.Count - 1);
                }
            }
            return Vlist[0];
        }
    
    }

    public struct FormulaItem
    {
        public string Name;
        public FormulaItemType Typ;
    }

    public enum FormulaItemType
    {
        Operator,
        Number,
        Tag,
        EndForma
    }
}
