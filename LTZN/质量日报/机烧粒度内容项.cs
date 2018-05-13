using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;

namespace LTZN.质量日报
{
    [DataObjectAttribute()]
    public  class 机烧粒度内容项
    {
        private 机烧粒度内容项数据 currentData;
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
        public DateTime SJ
        {
            get
            {
                return this.currentData.SJ;
            }
            set
            {
                this.currentData.SJ = value;
            }
        }
        [Bindable(true)]
        public double? ZG
        {
            get
            {
                return this.currentData.ZG;
            }
            set
            {
                this.currentData.ZG = value;
            }
        }
        [Bindable(true)]
        public double? KM
        {
            get
            {
                return this.currentData.KM;
            }
            set
            {
                this.currentData.KM = value;
            }
        }
        [Bindable(true)]
        public double? SF
        {
            get
            {
                return this.currentData.SF;
            }
            set
            {
                this.currentData.SF = value;
            }
        }
        [Bindable(true)]
        public double? KY
        {
            get
            {
                return this.currentData.KY;
            }
            set
            {
                this.currentData.KY = value;
            }
        }
        [Bindable(true)]
        public double? DY40
        {
            get
            {
                return this.currentData.DY40;
            }
            set
            {
                this.currentData.DY40 = value;
            }
        }
        [Bindable(true)]
        public double? ZJ425
        {
            get
            {
                return this.currentData.ZJ425;
            }
            set
            {
                this.currentData.ZJ425 = value;
            }
        }
        [Bindable(true)]
        public double? ZJ2516
        {
            get
            {
                return this.currentData.ZJ2516;
            }
            set
            {
                this.currentData.ZJ2516 = value;
            }
        }
        [Bindable(true)]
        public double? ZJ1610
        {
            get
            {
                return this.currentData.ZJ1610;
            }
            set
            {
                this.currentData.ZJ1610 = value;
            }
        }
        [Bindable(true)]
        public double? ZJ105
        {
            get
            {
                return this.currentData.ZJ105;
            }
            set
            {
                this.currentData.ZJ105 = value;
            }
        }
        [Bindable(true)]
        public double? XY5
        {
            get
            {
                return this.currentData.XY5;
            }
            set
            {
                this.currentData.XY5 = value;
            }
        }
        [Bindable(true)]
        public double? RI
        {
            get
            {
                return this.currentData.RI;
            }
            set
            {
                this.currentData.RI = value;
            }
        }
        [Bindable(true)]
        public double? TBS
        {
            get
            {
                return this.currentData.TBS;
            }
            set
            {
                this.currentData.TBS = value;
            }
        }
        [Bindable(true)]
        public double? TBE
        {
            get
            {
                return this.currentData.TBE;
            }
            set
            {
                this.currentData.TBE = value;
            }
        }
        [Bindable(true)]
        public double? TB
        {
            get
            {
                return this.currentData.TB;
            }
            set
            {
                this.currentData.TB = value;
            }
        }
        [Bindable(true)]
        public double? RDI63
        {
            get
            {
                return this.currentData.RDI63;
            }
            set
            {
                this.currentData.RDI63 = value;
            }
        }
        [Bindable(true)]
        public double? RDI315
        {
            get
            {
                return this.currentData.RDI315;
            }
            set
            {
                this.currentData.RDI315 = value;
            }
        }
        [Bindable(true)]
        public double? RDI05
        {
            get
            {
                return this.currentData.RDI05;
            }
            set
            {
                this.currentData.RDI05 = value;
            }
        }
       
         public 机烧粒度内容项(string MC, DateTime SJ)
        {
            this.MC = MC;
            this.SJ = SJ;
        }
         protected 机烧粒度内容项()
        { 
        }

    }
}
