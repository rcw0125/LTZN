using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Text;
using System.ComponentModel;
using LTZN.Data;
using System.Windows.Forms;

namespace LTZN.CalFramework
{
    public class CalTag : ObservableObject
    {
        private string _eid = "";
        /// <summary>
        /// 主键
        /// </summary>
        public string EID
        {
            get { return _eid; }
            set
            {
                if (_eid != value)
                {
                    _eid = value;
                    RaisePropertyChanged("EID");
                }
            }
        }

        private CalModel _myModel;
        /// <summary>
        /// 所属模型
        /// </summary>
        public CalModel MyModel
        {
            get { return _myModel; }
            set
            {
                if (_myModel != value)
                {
                    _myModel = value;
                    RaisePropertyChanged("MyModel");
                }
            }
        }


        private CalOrg _myCalOrg;
        /// <summary>
        /// 所属组织
        /// </summary>
        public CalOrg MyCalOrg
        {
            get { return _myCalOrg; }
            set
            {
                if (_myCalOrg != value)
                {
                    _myCalOrg = value;
                    RaisePropertyChanged("MyCalOrg");
                }
            }
        }

        private string _tagName = "";
        /// <summary>
        /// 标签名称
        /// </summary>
        public string TagName
        {
            get { return _tagName; }
            set
            {
                if (_tagName != value)
                {
                    _tagName = value;
                    RaisePropertyChanged("TagName");
                }
            }
        }

        public string TagFullName
        {
            get
            {
                if (this.MyCalOrg != null)
                    return this.MyCalOrg.OrgFullName + "." + TagName;
                else
                    return TagName;
            }
        }

        private string _forma = "";
        /// <remarks>公式</remarks>
        public string Forma
        {
            get { return _forma; }
            set
            {
                if (_forma != value)
                {
                    _forma = value;
                    RaisePropertyChanged("Forma");
                }
            }
        }

        private int? _dec;
        /// <summary>
        /// 小数位数
        /// </summary>
        public int? Dec
        {
            get { return _dec; }
            set
            {
                if (_dec != value)
                {
                    _dec = value;
                    RaisePropertyChanged("Dec");
                }
            }
        }

        public int Idx
        {
            get
            {
                if (this.MyCalOrg != null)
                {
                    return this.MyCalOrg.CalTags.IndexOf(this);
                }
                else
                    return 99999;
            }
        }

        private double _tagValue = 0;
        /// <remarks>值</remarks>
        public double TagValue
        {
            get
            {
                if (IsLock)
                    CalFinished = true;
                if (!CalFinished)
                    Cal();
                return _tagValue;
            }
            set
            {
                double temp = value;
                if (Dec.HasValue)
                    temp = Math.Round(value, Dec.Value);
                if (_tagValue != temp)
                {
                    _tagValue = temp;
                    RaisePropertyChanged("TagValue");
                }
            }
        }

        private bool _calFinished = true;
        /// <remarks>是否计算完成</remarks>
        public bool CalFinished
        {
            get { return _calFinished; }
            set
            {
                if (_calFinished != value)
                {
                    _calFinished = value;
                    RaisePropertyChanged("CalFinished");
                }
            }
        }

        private bool _IsLock=false;
        //是否锁定
        public bool IsLock
        {
            get { return _IsLock; }
            set
            {
                if (_IsLock != value)
                {
                    _IsLock = value;
                    RaisePropertyChanged("IsLock");
                }
            }
        }

        //自动设置
        public void SetAutoValue(double val)
        {
            if (!IsLock)
            {
                TagValue = val;
            }
        }

        //手动设置
        public void SetManualValue(double val)
        {
            TagValue = val;
        }
        //手动设置
        public void SetManualValue(double val, bool isLock)
        {
            TagValue = val;
            isLock = true;
        }

        public virtual void Cal()
        {
            if (!CalFinished && !IsLock)
            {
                if (!string.IsNullOrEmpty(Forma))
                { 
                    CalFinished = true;
                    this.TagValue = Convert.ToDouble(DoCal(Forma));// MyModel.DoForma(Forma);

                   
                }
            }
        }

        public List<CalTag> GetChildTags()
        {
            List<CalTag> subTags = new List<CalTag>();
            if (this.MyCalOrg != null)
            {
                foreach (var item in this.MyCalOrg.ChildCalOrgs)
                {
                    CalTag subTag = item.GetTag(this.TagName);
                    if (subTag != null)
                    {
                        subTags.Add(subTag);
                    }
                }
            }
            else if (this.MyModel != null)
            {
                foreach (var item in this.MyModel.RootOrgs)
                {
                    CalTag subTag = item.GetTag(this.TagName);
                    if (subTag != null)
                    {
                        subTags.Add(subTag);
                    }
                }
            }
            return subTags;
        }

        public CalTag GetBrotherTag(string tagName)
        {
            CalTag brotherTag = null;
            if (this.MyCalOrg != null)
            {
                brotherTag = this.MyCalOrg.GetTag(tagName);
            }

            if (brotherTag == null && this.MyModel != null)
            {
                brotherTag = this.MyModel.GetTag(tagName);
            }
            return brotherTag;
        }


        private OperFactory of = new OperFactory();
        /// <summary>
        /// 开始计算
        /// </summary>
        public object DoCal(string str)
        {
            /// 算术符栈
            ArrayList HList = new ArrayList();
            /// 数值栈
            ArrayList Vlist = new ArrayList();

            //读算试工具
            CalUtility cu = new CalUtility(str);

            FormaItem formaItem = cu.getItem();
            while (true)
            {
                if (formaItem.Typ == FormaItemType.Number)
                {
                    //如果是数值，则写入数据栈
                    Vlist.Add(formaItem.Name);
                }
                else if (formaItem.Typ == FormaItemType.Tag)
                {
                    double? tagvalue = GetTagValue(formaItem.Name);
                    if (tagvalue == null)
                        Vlist.Add("0");
                    else
                        Vlist.Add(tagvalue.Value.ToString());
                }
                else
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

                formaItem = cu.getItem();
            }
            return Vlist[0];
        }

        string tagvalue = "";
        public double? GetTagValue(string tagName)
        {
            tagvalue = tagName + "1";

            try
            {
                CalTag result = null;
                if (this.MyCalOrg != null)
                    result = MyCalOrg.GetTag(tagName);
                else if (MyModel != null)
                    result = MyModel.GetTag(tagName);
                tagvalue = tagName + "2";

                if (result != null)
                    return result.TagValue;
                else
                    return null;
            
            }
            catch (Exception ex)
            {
                MessageBox.Show(tagvalue + "标签计算数值出错" + ex.ToString());
                return 0;
            }
            
        }
    }
}
