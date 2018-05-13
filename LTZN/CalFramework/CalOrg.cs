using System;
using System.Collections.Generic;
using System.Text;
using LTZN.Data;

namespace LTZN.CalFramework
{
   /// <summary>
   /// 组织关系节点
   /// </summary>
    public class CalOrg : ObservableObject
    {
        private string _eid="";
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

        private string _orgName="";
        /// <summary>
        /// 组织名称
        /// </summary>
        public string OrgName
        {
            get { return _orgName; }
            set
            {
                if (_orgName != value)
                {
                    _orgName = value;
                    RaisePropertyChanged("EName");
                }
            }
        }

        public string OrgFullName
        {
            get {
                if (this.ParentCalOrg != null)
                    return this.ParentCalOrg.OrgFullName + "." + OrgName;
                else
                    return OrgName;
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


        private CalOrg _parentCalOrg;
        /// <summary>
        /// 上级组织
        /// </summary>
        public CalOrg ParentCalOrg
        {
            get { return _parentCalOrg; }
            set
            {
                if (_parentCalOrg != value)
                {
                    _parentCalOrg = value;
                    RaisePropertyChanged("ParentCalOrg");
                }
            }
        }


        private List<CalOrg> _childCalOrgs=new List<CalOrg>();
        /// <summary>
        /// 下级组织
        /// </summary>
        public List<CalOrg> ChildCalOrgs
        {
            get { return _childCalOrgs; }
        }


        private List<CalTag> _calTags=new List<CalTag>();
        /// <summary>
        /// 所含计算标签
        /// </summary>
        public List<CalTag> CalTags
        {
            get { return _calTags; }
        }


        public void UpIdx(CalTag calTag)
        {
            int idx = _calTags.IndexOf(calTag);
            if (idx > 0)
            {
                _calTags.Remove(calTag);
                _calTags.Insert(idx - 1, calTag);
            }
        }
        public void DownIdx(CalTag calTag)
        {
            int idx = _calTags.IndexOf(calTag);
            if (idx >= 0 && idx < _calTags.Count - 1)
            {
                _calTags.Remove(calTag);
                _calTags.Insert(idx + 1, calTag);
            }
        }

        public CalTag GetTag(string tagName)
        {
           
            if (string.IsNullOrEmpty(tagName)) return null;

            CalTag result = null;

            string[] spilts = tagName.Split('.');

            int splitIdx = tagName.IndexOf('.');

            bool findTag = false;
            string subTagName = tagName;
            if (splitIdx == -1) //如果找不到分隔符表明这是一个标签
            {
                foreach (var item in CalTags)
                {
                    if (item.TagName == tagName)
                    {
                        result = item;
                        findTag = true;
                        break;
                    }
                }
            }
            else if (splitIdx > 0)
            {
                string subOrgName = tagName.Substring(0, splitIdx);
                foreach (var subOrg in ChildCalOrgs)
                {
                    if (subOrg.OrgName == subOrgName)
                    {   
                        subTagName = tagName.Substring(splitIdx + 1, tagName.Length - splitIdx - 1);
                        result = subOrg.GetTag(subTagName);
                        findTag = true;
                        break;
                    }
                }
            }
            else //第一个字符是 .  或 没有找到下级目录 到上级目录查找
            {
                subTagName = tagName.Substring(1, tagName.Length - 1);
            }
            if (!findTag)
            {
                if (this.ParentCalOrg != null)
                    result = ParentCalOrg.GetTag(subTagName);
                else if (MyModel != null)
                    result = MyModel.GetTag(subTagName);
            }

            return result;
        }

        public CalOrg AddOrg(string orgName)
        {
            foreach (var item in ChildCalOrgs)
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
            org.MyModel = this.MyModel;
            org.ParentCalOrg = this;
            org.ParentCalOrg.ChildCalOrgs.Add(this);
            this.MyModel.AllCalOrgs.Add(org);
            return org;
        }

        public CalTag AddTag(string tagName)
        {
            foreach (var item in CalTags)
            {
                if (item.TagName == tagName)
                {
                    System.Windows.Forms.MessageBox.Show("相同目录下已有此标签:" + tagName);
                    return item;
                }
            }
            CalTag tag = new CalTag();
            tag.EID = Guid.NewGuid().ToString();
            tag.TagName = tagName;
            tag.MyModel = this.MyModel;
            tag.MyCalOrg = this;
            this.CalTags.Add(tag);
            tag.MyModel.AllCalTags.Add(tag);
            return tag;
        }

        public void AddTagGroup(string tagName)
        {
            this.AddTag(tagName);
            foreach (var org in ChildCalOrgs)
            {
                org.AddTagGroup(tagName);   
            }
        }

        public void RemoveOrg(CalOrg org)
        {
            this.ChildCalOrgs.Remove(org);
            org.MyModel.AllCalOrgs.Remove(org);
        }

        public void RemoveTag(CalTag tag)
        {
            this.CalTags.Remove(tag);
            tag.MyModel.AllCalTags.Remove(tag);
        }
    }
}
