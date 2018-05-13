using System;
using System.Collections.Generic;
using System.Text;

namespace ZHCDB
{
    public interface ITagSource
    {
         double? GetDoubleValue(string TagName);

         string GetStringValue(string TagName);
    }
}
