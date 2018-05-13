using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;


namespace LTZN.质量日报
{
    [DataObjectAttribute()]
  public  class 质量日报内容项
    {

        public bool[] p = new bool[24];

        private 质量日报内容项数据 currentData;
       
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
        public double? TFE
        {
            get
            {
                return this.currentData.TFE;
            }
            set
            {
                this.currentData.TFE = value;
            }
        }
        [Bindable(true)]
        public double? FeO
        {
            get
            {
                return this.currentData.FeO;
            }
            set
            {
                this.currentData.FeO = value;
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
        public double? TiO2
        {
            get
            {
                return this.currentData.TiO2;
            }
            set
            {
                this.currentData.TiO2 = value;
            }
        }
        [Bindable(true)]
        public double? R
        {
            get
            {
                return this.currentData.R;
            }
            set
            {
                this.currentData.R = value;
            }
        }
        [Bindable(true)]
        public double? Cr
        {
            get
            {
                return this.currentData.Cr;
            }
            set
            {
                this.currentData.Cr = value;
            }
        }
        [Bindable(true)]
        public double? Ni
        {
            get
            {
                return this.currentData.Ni;
            }
            set
            {
                this.currentData.Ni = value;
            }
        }
        [Bindable(true)]
        public double? As
        {
            get
            {
                return this.currentData.As;
            }
            set
            {
                this.currentData.As = value;
            }
        }
        [Bindable(true)]
        public double? MnO
        {
            get
            {
                return this.currentData.MnO;
            }
            set
            {
                this.currentData.MnO = value;
            }
        }
        [Bindable(true)]
        public double? Pb
        {
            get
            {
                return this.currentData.Pb;
            }
            set
            {
                this.currentData.Pb = value;
            }
        }
        [Bindable(true)]
        public double? Zn
        {
            get
            {
                return this.currentData.Zn;
            }
            set
            {
                this.currentData.Zn = value;
            }
        }
        [Bindable(true)]
        public double? JJS
        {
            get
            {
                return this.currentData.JJS;
            }
            set
            {
                this.currentData.JJS = value;
            }
        }
        [Bindable(true)]
        public double? V2O5
        {
            get
            {
                return this.currentData.V2O5;
            }
            set
            {
                this.currentData.V2O5 = value;
            }
        }
        [Bindable(true)]
        public double? Cu
        {
            get
            {
                return this.currentData.Cu;
            }
            set
            {
                this.currentData.Cu = value;
            }
        }
        [Bindable(true)]
        public double? Mo
        {
            get
            {
                return this.currentData.Mo;
            }
            set
            {
                this.currentData.Mo = value;
            }
        }
        [Bindable(true)]
        public double? Sn
        {
            get
            {
                return this.currentData.Sn;
            }
            set
            {
                this.currentData.Sn = value;
            }
        }
        [Bindable(true)]
        public double? Sb
        {
            get
            {
                return this.currentData.Sb;
            }
            set
            {
                this.currentData.Sn = value;
            }
        }
         public 质量日报内容项(string MC)
        {
            this.MC = MC;
           
        }
        protected 质量日报内容项()
        { 
        }
    }
      
}
