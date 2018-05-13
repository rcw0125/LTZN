using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;

namespace LTZN.CalFramework2
{
    public class MemTagValChangedArgs : EventArgs
    {
        private MemTag sourceTag = null;
        private List<MemTag> _ChgMemTagList = new List<MemTag>();

        public MemTagValChangedArgs(MemTag tag)
        {
            sourceTag = tag;
            AddMemTag(tag);
        }

        /// <summary>
        /// 事件源标签
        /// </summary>
        public MemTag SourceTag
        {
            get
            {
                return sourceTag;
            }
        }
        /// <summary>
        /// 变动的标签列表
        /// </summary>
        public List<MemTag> ChgMemTagList
        {
            get
            {
                return _ChgMemTagList;
            }
        }


        /// <summary>
        /// 添加变动的标签
        /// </summary>
        /// <param name="tag"></param>
        public bool AddMemTag(MemTag tag)
        {
            if (!_ChgMemTagList.Contains(tag))
            {
                _ChgMemTagList.Add(tag);
                return true;
            }
            return false;
        }
    }
    
    public class MemTag : INotifyPropertyChanged
    {   
        /// <summary>
        /// 所属模块
        /// </summary>
        private MemTagModel memTagModel = null;
        
        private MemTagType _tagType = MemTagType.Numeric;
        /// <summary>
        /// 标签类型
        /// </summary>
        public MemTagType TagType
        {
            get
            {
                return _tagType;
            }
            set
            {
                _tagType = value;
            }
        }

        private string tag_name = "";
        /// <summary>
        /// 标签名
        /// </summary>
        public string TagName
        {
            get
            {
                return tag_name;
            }
            set
            {
                tag_name = value;
            }
        }
 
        private string _strVal = "";
        private string db_strVal = "";
        /// <summary>
        /// 标签字符串值
        /// </summary>
        public string StrVal
        {
            get
            {
                return _strVal;
            }
            set
            {
                if (_strVal != value)
                {
                    _strVal = value;
                    this.memTagModel.NotifyTagValueChanged(this);
                }
            }
        }

        private double _tagVal = 0;
        private double db_tagVal = 0;
        /// <summary>
        /// 标签数值
        /// </summary>
        public double TagVal
        {
            get
            {
                if (NeedCal && cu!=null)
                {
                    //double? temp = Cal(this.tag_name,this.Forma);
                    double? temp = Convert.ToDouble(cu.DoCal(this.memTagModel));

                    if (Dec.HasValue)
                        _tagVal = Math.Round(temp ?? 0, Dec.Value);
                    else
                        _tagVal = temp ?? 0;
                    needcal = false;
                }
                return _tagVal;
            }
            set
            {
                double temp;
                if (Dec.HasValue)
                    temp = Math.Round(value, Dec.Value);
                else
                    temp = value;
                if (temp != _tagVal)
                {
                    _tagVal = temp;
                    ReportChanged();
                }
            }
        }

       /// <summary>
       /// 标签精度
       /// </summary>
        public int? Dec { get; set; }

        private CalUtility cu = null;
        private string _forma = "";
        /// <summary>
        /// 计算公式
        /// </summary>
        public string Formula
        {
            get { return _forma; }
            set
            {
                if (_forma != value)
                {
                    _forma = value;
                    cu = new CalUtility(_forma);
                    this.memTagModel.UnRegister(this);
                    foreach (var item in cu)
                    {
                        if (item.Typ == FormulaItemType.Tag)
                        {
                            this.memTagModel.reg(item.Name, this);
                        }
                    }
                    RaisePropertyChanged("Forma");
                }
            }
        }

       /// <summary>
       /// 是否锁定
       /// </summary>
        public bool IsLock { get; set; }
        //public bool NeedInsert { get; set; }

        private bool needcal = true;
        /// <summary>
        /// 是否需重新计算
        /// </summary>
        public bool NeedCal
        {
            get
            {
                if (!string.IsNullOrEmpty(this.Formula))
                    return needcal;
                else
                    return false;
            }
            set
            {
                if (needcal != value)
                {
                    needcal = value;
                }
            }
        }

        public MemTag(MemTagModel memTagModel, string tagname, int? dec, string formula, bool needSaveToDb)
        {
            this.memTagModel = memTagModel;
            this.TagType = MemTagType.Numeric;
            this.TagName = tagname;
            this.Formula = formula;
            this.Dec = dec;
        }

        public MemTag(MemTagModel memTagModel, string tagname,string strval,bool needSaveToDb)
        {
            this.memTagModel = memTagModel;
            this.TagType = MemTagType.String;
            this.TagName = tagname;
        }

        private void ReportChanged()
        {
            this.memTagModel.NotifyTagValueChanged(this);
            RaisePropertyChanged("TagVal");
        }

        public void NotifyTagValueChanged(MemTagValChangedArgs arg)
        {
            if (!this.needcal)
            {
                needcal = true;
                this.memTagModel.NotifyTagValueChanged(arg, TagName);
            }
        }

        #region INotifyPropertyChanged 成员

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void RaisePropertyChanged(string propertyName)
        {
            var handler = PropertyChanged;

            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        #endregion

        //-------数据持久化相关
        private bool _ExistsInDb = false;
        private bool _NeedSaveToDb = true;

        /// <summary>
        /// 是否存在于数据库
        /// </summary>
        public bool ExistsInDb
        {
            get
            {
                return _ExistsInDb;
            }
        }

        /// <summary>
        /// 是否需要存入数据库
        /// </summary>
        public bool NeedSaveToDb
        {
            get
            {
                return _NeedSaveToDb;
            }
        }

        public void DbInit(bool indb)
        {
            db_strVal = StrVal;
            db_tagVal = TagVal;
            _ExistsInDb = indb;
        }

        public DataStateType DataState
        {
            get
            {
                if (NeedSaveToDb)
                {
                    if (ExistsInDb)
                    {
                        if (TagType == MemTagType.String)
                        {
                            if (string.IsNullOrEmpty(StrVal)) return DataStateType.Delete;
                            if (StrVal != db_strVal)
                            {
                                return DataStateType.Update;
                            }
                        }
                        else
                        {
                            if (TagVal == 0 || double.IsNaN(TagVal) || double.IsInfinity(TagVal)) return DataStateType.Delete;

                            if (TagVal != db_tagVal)
                            {
                                return DataStateType.Update;
                            }
                        }
                        return DataStateType.UnChanged;
                    }
                    else
                    {
                        if (TagType == MemTagType.String)
                        {
                            if (string.IsNullOrEmpty(StrVal)) return DataStateType.UnChanged;
                            if (StrVal != db_strVal)
                            {
                                return DataStateType.Add;
                            }
                            else
                            {
                                return DataStateType.UnChanged;
                            }
                        }
                        else
                        {
                            if (TagVal == 0 || double.IsNaN(TagVal) || double.IsInfinity(TagVal))
                            {
                                return DataStateType.UnChanged;
                            }
                            else
                            {
                                return DataStateType.Add;
                            }
                        }
                    }
                }
                else
                {
                    if (ExistsInDb)
                        return DataStateType.Delete;
                    else
                        return DataStateType.UnChanged;
                }
            }
        }

        //-------------TableViews
        Dictionary<string, List<System.Drawing.Point>> viewPosTable = new Dictionary<string, List<System.Drawing.Point>>();
        public void AddViewPoint(string viewName, System.Drawing.Point point)
        {
            if (!viewPosTable.ContainsKey(viewName))
                viewPosTable.Add(viewName, new List<System.Drawing.Point>());
            if (!viewPosTable[viewName].Contains(point))
                viewPosTable[viewName].Add(point);
        }
        public List<System.Drawing.Point> GetViewPoint(string viewName)
        {
            if (viewName == "" && viewPosTable.Count > 0)
            {
                Dictionary<string, List<System.Drawing.Point>>.Enumerator it = viewPosTable.GetEnumerator();
                it.MoveNext();
                return it.Current.Value;
            }
            else if (viewPosTable.ContainsKey(viewName))
                return viewPosTable[viewName];
            else
                return new List<System.Drawing.Point>();
        }

        public void SetVal(double val)
        {
            TagVal = val;
        }
        public void SetVal(string val)
        {
            StrVal = "";
        }
        public void SetInitVal(double val)
        {
            _tagVal = val;
        }
        public void SetInitVal(string val)
        {
            _strVal = "";
        }
        public void SetDbVal(double val)
        {
            _tagVal = val;
            _ExistsInDb = true;
        }
        public void SetDbVal(string val)
        {
            _strVal = val;
            _ExistsInDb = true;
        }
        public void clear()
        {
            _tagVal = 0;
            _strVal = "";
            _ExistsInDb = false;
            needcal = true;
        }
    }

    public enum MemTagType
    {
        Numeric,
        String
    }
}
