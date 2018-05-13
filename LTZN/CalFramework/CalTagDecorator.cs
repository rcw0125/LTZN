using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.CalFramework
{
   public abstract  class CalTagDecorator:CalTag 
    {
       protected CalTag component = null;

       //public string TagName
       //{
       //    get
       //    {
       //        return component.TagName;
       //    }
       //    set
       //    {
       //        component.TagName = value;
       //    }
       //}

       //public double TagValue
       //{
       //    get
       //    {
       //        return component.TagValue;
       //    }
       //    set
       //    {
       //        component.TagValue = value;
       //    }
       //}

       //public bool CalFinished
       //{
       //    get
       //    {
       //        return component.CalFinished;
       //    }
       //    set
       //    {
       //        component.CalFinished = value;
       //    }
       //}

       public CalTagDecorator(CalTag tag)
       {
           this.component = tag;
       }

    }
}
