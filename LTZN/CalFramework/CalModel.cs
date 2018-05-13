using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Text;
using LTZN.Data;
using System.Windows.Forms;

namespace LTZN.CalFramework
{
    public class CalModel : ObservableObject
    {

        public CalModel()
        {
        }

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

        private string _modelName = "";
        /// <summary>
        /// 模型名称
        /// </summary>
        public string ModelName
        {
            get { return _modelName; }
            set
            {
                if (_modelName != value)
                {
                    _modelName = value;
                    RaisePropertyChanged("EName");
                }
            }
        }


        private List<CalOrg> _rootOrgs=new List<CalOrg>();
        /// <summary>
        /// 模型根级组织
        /// </summary>
        public List<CalOrg> RootOrgs
        {
            get { return _rootOrgs; }
        }


        private List<CalTag> _rootTags=new List<CalTag>();
        /// <summary>
        /// 模型根级计算点
        /// </summary>
        public List<CalTag> RootTags
        {
            get { return _rootTags; }
        }


        private List<CalOrg> _allCalOrgs=new List<CalOrg>();
        /// <summary>
        /// 所有组织
        /// </summary>
        public List<CalOrg> AllCalOrgs
        {
            get { return _allCalOrgs; }
        }

        private List<CalTag> _allCalTags=new List<CalTag>();
        /// <summary>
        /// 所有计算标签
        /// </summary>
        public List<CalTag> AllCalTags
        {
            get { return _allCalTags; }
        }

        public void Init()
        {
            foreach (CalTag item in AllCalTags)
            {
                item.TagValue = 0;
                item.CalFinished = false;
            }
        }

        public CalOrg AddOrg(string orgName)
        {
            foreach (var item in RootOrgs)
            {
                if (item.OrgName == orgName)
                {
                    System.Windows.Forms.MessageBox.Show("相同目录下已有此组织:" + orgName);
                    break;
                }
            }
            CalOrg org = new CalOrg();
            org.EID = Guid.NewGuid().ToString();
            org.OrgName = orgName;
            org.MyModel = this;
            AllCalOrgs.Add(org);
            RootOrgs.Add(org);
            return org;
        }

        public CalTag AddTag(string tagName)
        {
            foreach (var item in RootTags)
            {
                if (item.TagName == tagName)
                {
                    System.Windows.Forms.MessageBox.Show("相同目录下已有此标签:" + tagName);
                    break;
                }
            }
            CalTag tag = new CalTag();
            tag.EID = Guid.NewGuid().ToString();
            tag.TagName = tagName;
            tag.MyModel = this;
            AllCalTags.Add(tag);
            RootTags.Add(tag);
            return tag;
        }

        public void AddTagGroup(string tagName)
        {
            foreach (CalOrg  org in AllCalOrgs)
            {
                org.AddTag(tagName);
            }
        }

        public void RemoveOrg(CalOrg org)
        {
            RootOrgs.Remove(org);
            AllCalOrgs.Remove(org);
        }

        public void RemoveTag(CalTag tag)
        {
            RootTags.Remove(tag);
            AllCalTags.Remove(tag);
        }

        public CalTag GetTag(string tagName)
        {
            string shortTagName = tagName;
            CalOrg myOrg = null;
            if (tagName.Contains("."))
            {
                string[] items = tagName.Split('.');
                for (int i = 0; i < items.Length - 1; i++)
                {
                    if (myOrg == null)
                    {
                        foreach (var org in this.RootOrgs)
                        {
                            if (org.OrgName == items[i])
                            {
                                myOrg = org;
                                break;
                            }
                        }
                    }
                    else
                    {
                        foreach (var org in myOrg.ChildCalOrgs)
                        {
                            if (org.OrgName == items[i])
                            {
                                myOrg = org;
                                break;
                            }
                        }
                    }
                }
                 shortTagName = items[items.Length - 1];
            }
            CalTag result = null;
            if (myOrg == null)
            {
                foreach (var item in this.RootTags)
                {
                    if (item.TagName == shortTagName)
                    {
                        result = item;
                        break;
                    }
                }
                //if (result == null)
                //    result = this.AddTag(shortTagName);
            }
            else
            {
                foreach (var item in myOrg.CalTags)
                {
                    if (item.TagName == shortTagName)
                    {
                        result = item;
                        break;
                    }
                }
                //if (result == null)
                //    result = myOrg.AddTag(shortTagName);
            }

            return result;
        }

        string tagname = "";
        public void Calc()
        {
            try
            {
                foreach (CalTag tag in AllCalTags)
                {
                    if (!string.IsNullOrEmpty(tag.Forma))
                        tag.CalFinished = false;

                }

                foreach (CalTag tag in AllCalTags)
                {
                    tagname = tag.TagName+"9";
                    tag.Cal();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(tagname+"--calmodel出错" + ex.ToString());
                return;
            }
          
        }

        public double GetTagValue(string tagName)
        {
            CalTag tag = GetTag(tagName);
            if (tag != null)
            { 
                 return tag.TagValue;
            }
            else
                return 0;
        }

        public double? GetNullTagValue(string tagName)
        {
            CalTag tag = GetTag(tagName);
            if (tag != null)
            {
                double result = tag.TagValue;
                if (result == 0)
                    return null;
                else
                    return result;
            }
            else
                return null;
        }

        public double DoForma(string forma)
        {
            return Convert.ToDouble(DoCal(forma));
        }

        #region 计算部分
        
        /// <summary>
        /// 运算操作器工厂
        /// </summary>
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
                    Vlist.Add(GetTagValue(formaItem.Name).ToString());
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
       
        #endregion

    }

}
