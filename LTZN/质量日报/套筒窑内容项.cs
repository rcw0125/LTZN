using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;

namespace LTZN.质量日报
{
 
    [DataObjectAttribute()]
   
      
       
    public class 套筒窑内容项
    {

           private 套筒窑内容项数据 currentData;


           [Bindable(true)]
           public string MC
           {
               get
               {
                   return this.currentData.MC;
               }
               set
               {
                   this.currentData.MC = value;
               }
           }
           [Bindable(true)]
           public double? CaO
           {
               get
               {
                   return this.currentData.CaO;
               }
               set
               {
                   this.currentData.CaO = value;
               }
           }
           [Bindable(true)]
           public double? SiO2
           {
               get
               {
                   return this.currentData.SiO2;
               }
               set
               {
                   this.currentData.SiO2 = value;
               }
           }
           [Bindable(true)]
           public double? Al2O3
           {
               get
               {
                   return this.currentData.Al2O3;
               }
               set
               {
                   this.currentData.Al2O3 = value;
               }
           }
           [Bindable(true)]
           public double? MgO
           {
               get
               {
                   return this.currentData.MgO;
               }
               set
               {
                   this.currentData.MgO = value;
               }
           }
          [Bindable(true)]
        public double? S
        {
            get
            {
                return this.currentData.S;
            }
            set
            {
                this.currentData.S = value;
            }
        }
        [Bindable(true)]
        public double? P
        {
            get
            {
                return this.currentData.P;
            }
            set
            {
                this.currentData.P = value;
            }
        }
        [Bindable(true)]
        public double? HXD
        {
            get
            {
                return this.currentData.HXD;
            }
            set
            {
                this.currentData.HXD = value;
            }
        }
        [Bindable(true)]
        public double? ZJ
        {
            get
            {
                return this.currentData.ZJ;
            }
            set
            {
                this.currentData.ZJ = value;
            }
        }
    
        [Bindable(true)]
        public double? DY60
        {
            get
            {
                return this.currentData.DY60;
            }
            set
            {
                this.currentData.DY60 = value;
            }
        }
        [Bindable(true)]
        public double? ZJ64
        {
            get
            {
                return this.currentData.ZJ64;
            }
            set
            {
                this.currentData.ZJ64 = value;
            }
        }
        [Bindable(true)]
        public double? XY40
        {
            get
            {
                return this.currentData.XY40;
            }
            set
            {
                this.currentData.XY40 = value;
            }
        }

    }
}
