using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace LTZN.CalFramework
{
        /// <summary>
        /// Calculate 的摘要说明。
        /// 计算实现主类
        /// </summary>
        public class Calculate
        {

            private CalModel model = null;

            /// <summary>
            /// 算术符栈
            /// </summary>
            private ArrayList HList;
            /// <summary>
            /// 数值栈
            /// </summary>
            public ArrayList Vlist;
            /// <summary>
            /// 读算试工具
            /// </summary>
            private CalUtility cu;
            /// <summary>
            /// 运算操作器工厂
            /// </summary>
            private OperFactory of;
            /// <summary>
            /// 构造方法
            /// </summary>
            /// <param name="str">算式</param>
            public Calculate(CalModel model)
            {
                this.model = model;
                HList = new ArrayList();
                Vlist = new ArrayList();
                of = new OperFactory();
               // cu = new CalUtility(str);
            }
           

            /// <summary>
            /// 开始计算
            /// </summary>
            public object DoCal(string str)
            {
                HList.Clear();
                Vlist.Clear();
                cu = new CalUtility(str);

                FormaItem formaItem = cu.getItem();
                while (true)
                {
                   
                    if (formaItem.Typ==FormaItemType.Number)
                    {
                        //如果是数值，则写入数据栈
                        Vlist.Add(formaItem.Name);
                    }
                    else if (formaItem.Typ==FormaItemType.Operator)
                    {
                        //数值
                        Cal(formaItem.Name);
                    }
                    else if (formaItem.Typ == FormaItemType.Tag)
                    {
                        Vlist.Add(model.GetTagValue(formaItem.Name).ToString());
                    }
                    else
                    {
                        Cal("");
                        break;
                    }

                    formaItem = cu.getItem();
                }
                return Vlist[0];
            }
            /// <summary>
            /// 计算
            /// </summary>
            /// <param name="str">计算符</param>
            private void Cal(string str)
            {
                //符号表为空，而且当前符号为""，则认为已经计算完毕
                if (str.Equals("") && HList.Count == 0)
                    return;
                if (HList.Count > 0)
                {
                    //符号是否可以对消？
                    if (HList[HList.Count - 1].ToString().Equals("(") && str.Equals(")"))
                    {
                        HList.RemoveAt(HList.Count - 1);
                        if (HList.Count > 0)
                        {
                            str = HList[HList.Count - 1].ToString();
                            HList.RemoveAt(HList.Count - 1);
                            Cal(str);
                        }
                        return;
                    }
                    //比较优先级
                    if (cu.Compare(HList[HList.Count - 1].ToString(), str))
                    {
                        //如果优先，则计算
                        IOper p = of.CreateOper(HList[HList.Count - 1].ToString());
                        if (p != null)
                        {
                            Vlist[Vlist.Count - 2] = p.Oper(Vlist[Vlist.Count - 2], Vlist[Vlist.Count - 1]);
                            HList.RemoveAt(HList.Count - 1);
                            Vlist.RemoveAt(Vlist.Count - 1);
                            Cal(str);
                        }
                        return;
                    }
                    if (!str.Equals(""))
                        HList.Add(str);
                }
                else
                {
                    if (!str.Equals(""))
                        HList.Add(str);
                }
            }
        }
    }
