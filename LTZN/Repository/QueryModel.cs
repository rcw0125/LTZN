using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text;
using LTZN.CalFramework;

namespace LTZN.Repository
{
    public class QueryModel : RepositoryBase
    { 
        public void SetValue(string tagName, double val)
        {
            if (myModel != null)
            {
                CalTag tag = myModel.GetTag(tagName);
                if (tag == null)
                {
                    string err = "找不到标签：" + tagName;
                    if (!_errs.Contains(err))
                        _errs.Add(err);
                }
                else
                    tag.TagValue = val;
            }
        }

        public CalModel myModel = null;

        public QueryModel(string modelName)
        {
            CalModelRepository rp = new CalModelRepository();
            myModel = rp.GetModel(modelName);
            if (myModel == null)
                _errs.Add("找不到计算模型“" + modelName + "”");
        }

        public void ModelCalc()
        {
            if (this.myModel != null)
            {
                this.myModel.Calc();
            }
        }

        public void display(HtmlDocument doc)
        {
            if (doc != null)
            {
                doc.OpenNew(true);

                if (this.myModel != null)
                {
                    string content = Properties.Resources.ResourceManager.GetString(myModel.ModelName);
                    foreach (var tag in this.myModel.AllCalTags)
                    {
                        if (tag.TagValue == 0)
                            content = content.Replace("$" + tag.TagFullName + "$", "&nbsp;");
                        else
                            content = content.Replace("$" + tag.TagFullName + "$", tag.TagValue.ToString());
                    } doc.Write(content);
                }
                foreach (var item in Errs)
                {
                    doc.Write("<br>" + item + "</br>");
                }
            }
        }

        private  List<string> _errs = new List<string>();
        public List<string> Errs
        {
            get
            {
                return _errs;
            }
        }
    }
}
