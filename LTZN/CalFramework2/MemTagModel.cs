using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;

namespace LTZN.CalFramework2
{

    public class MemTagModel : IEnumerable<MemTag>
    {
        protected Dictionary<string, MemTag> alltag = new Dictionary<string, MemTag>();
        private Dictionary<string, Dictionary<int, string>> views = new Dictionary<string, Dictionary<int, string>>();
        private Dictionary<string, List<MemTag>> calRelation = new Dictionary<string, List<MemTag>>();

        public void UnRegister(MemTag targetTag)
        {
            foreach (var item in calRelation.Values)
            {
                if (item.Contains(targetTag))
                {
                    item.Remove(targetTag);
                }
            }
        }

        public void reg(string sourceTag, MemTag targetTag)
        {
            if (calRelation.ContainsKey(sourceTag))
            {
                calRelation[sourceTag].Add(targetTag);
            }
            else
            {
                List<MemTag> tags = new List<MemTag>();
                calRelation.Add(sourceTag, tags);
                tags.Add(targetTag);
            }
        }

        public void NotifyTagValueChanged(MemTag sourceTag)
        {
            MemTagValChangedArgs arg = new MemTagValChangedArgs(sourceTag);
            sourceTag.NotifyTagValueChanged(arg);
            if (TagValueChanged != null)
            {
                TagValueChanged(arg);
            }
        }

        public void NotifyTagValueChanged(MemTagValChangedArgs arg,string tagName)
        {
            if (calRelation.ContainsKey(tagName))
            {
                List<MemTag> tags = calRelation[tagName];
                foreach (var item in tags)
                {
                    item.NotifyTagValueChanged(arg);
                }
            }
        }
      
        /// <summary>
        /// 标签值改变事件
        /// </summary>
        public event Action<MemTagValChangedArgs> TagValueChanged;

        /// <summary>
        /// 添加标签
        /// </summary>
        /// <param name="tagname">标签名</param>
        /// <param name="forma">计算公式</param>
        /// <param name="dec">小数点位数</param>
        /// <param name="defaultVal">默认值</param>
        /// <param name="savetodb">是否保存到数据库</param>
        public MemTag AddNumericTag(string tagname, string forma, int? dec, double? defaultVal, bool savetodb)
        {
            MemTag tag = new MemTag(this, tagname, dec, forma,savetodb);
            alltag.Add(tagname, tag);
            return tag;
        }


        public MemTag AddStringTag(string tagname, bool saveToDb)
        {
            MemTag tag = new MemTag(this, tagname,"",saveToDb);
            alltag.Add(tagname, tag);
            return tag;
        }

        /// <summary>
        /// 得到标签
        /// </summary>
        /// <param name="tagname"></param>
        /// <returns></returns>
        public MemTag GetTag(string tagname)
        {
            if (alltag.ContainsKey(tagname))
                return alltag[tagname];
            else
                return null;
        }

        /// <summary>
        /// 获取标签值
        /// </summary>
        /// <param name="tagname">标签名</param>
        /// <returns>标签值</returns>
        public double? GetTagValue(string tagname)
        {
            if (alltag.ContainsKey(tagname))
                return alltag[tagname].TagVal;
            else
                return null;
        }

        /// <summary>
        /// 设置标签值
        /// </summary>
        /// <param name="tagname">标签名</param>
        /// <param name="val">标签值</param>
        public void SetTagValue(string tagname, double val)
        {
            if (alltag.ContainsKey(tagname))
                alltag[tagname].TagVal = val;
        }

        /// <summary>
        /// 设置标签值
        /// </summary>
        /// <param name="tagname">标签名</param>
        /// <param name="val">标签值</param>
        public void SetTagValue(string tagname, string strval)
        {
            if (alltag.ContainsKey(tagname))
            {
                alltag[tagname].StrVal = strval;
            }
        }

        /// <summary>
        /// 设置初始值（不引发数值变动事件)
        /// </summary>
        /// <param name="tagname">标签名</param>
        /// <param name="val">标签值</param>
        public void SetInitValue(string tagname, double val)
        {
            if (alltag.ContainsKey(tagname))
            {
                MemTag tag = alltag[tagname];
                tag.SetInitVal(val);
            }
        }


        /// <summary>
        /// 设置初始值（不引发数值变动事件)
        /// </summary>
        /// <param name="tagname">标签名</param>
        /// <param name="val">标签值</param>
        public void SetInitValue(string tagname, string val)
        {
            if (alltag.ContainsKey(tagname))
            {
                MemTag tag = alltag[tagname];
                tag.SetInitVal(val);
            }
        }


        /// <summary>
        /// 设置数据库值
        /// </summary>
        /// <param name="tagname">标签名</param>
        /// <param name="val">标签值</param>
        public void SetDbValue(string tagname, double val)
        {
            if (alltag.ContainsKey(tagname))
            {
                MemTag tag = alltag[tagname];
                tag.SetDbVal(val);
            }
        }


        /// <summary>
        /// 设置数据库值
        /// </summary>
        /// <param name="tagname">标签名</param>
        /// <param name="val">标签值</param>
        public void SetDbValue(string tagname, string val)
        {
            if (alltag.ContainsKey(tagname))
            {
                MemTag tag = alltag[tagname];
                tag.SetDbVal(val);

            }
        }

        /// <summary>
        /// 得到标签
        /// </summary>
        /// <param name="viewname">视图</param>
        /// <param name="row">行</param>
        /// <param name="col">列</param>
        /// <returns></returns>
        public virtual MemTag GetMemTag(string viewname, int row, int col)
        {
            if (!views.ContainsKey(viewname)) return null;
            int key = row * 1000 + col;
            if (!views[viewname].ContainsKey(key)) return null;
            string tagname = views[viewname][key];
            return GetTag(tagname);
        }

        /// <summary>
        /// 获取标签值
        /// </summary>
        /// <param name="viewname"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public double? GetTagValue(string viewname, int row, int col)
        {
            MemTag tag = GetMemTag(viewname, row, col);
            if (tag != null)
                return tag.TagVal;
            else
                return null;
        }

        /// <summary>
        /// 设置标签值
        /// </summary>
        /// <param name="viewname"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="val"></param>
        public virtual void SetTagValue(string viewname, int row, int col, double val)
        {
            MemTag tag = GetMemTag(viewname, row, col);
            if (tag != null)
                tag.TagVal = val;
        }

        /// <summary>
        /// 设置标签值
        /// </summary>
        /// <param name="viewname"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="val"></param>
        public virtual void SetTagValue(string viewname, int row, int col, string strval)
        {
            MemTag tag = GetMemTag(viewname, row, col);
            if (tag != null)
                tag.StrVal = strval;
        }

        /// <summary>
        /// 添加视图位置
        /// </summary>
        /// <param name="viewname"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="tagname"></param>
        public virtual void AddViewItem(string viewname, int row, int col, string tagname)
        {
            if (!views.ContainsKey(viewname))
                views.Add(viewname, new Dictionary<int, string>());
            views[viewname].Add(row * 1000 + col, tagname);
            GetTag(tagname).AddViewPoint(viewname, new System.Drawing.Point(row, col));
        }


        public void ReCal()
        {
            foreach (var item in alltag)
            {
                item.Value.NeedCal = true;
            }
        }


        #region IEnumerable<MemTag> 成员

        public IEnumerator<MemTag> GetEnumerator()
        {
            foreach (var item in alltag)
            {
                yield return item.Value;
            }
        }

        #endregion

        #region IEnumerable 成员

        IEnumerator IEnumerable.GetEnumerator()
        {
            foreach (var item in alltag)
            {
                yield return item.Value;
            }
        }

        #endregion
    }
}
