using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;
namespace LTZN.质量日报
{
    [DataObjectAttribute()]
   public class 焦炭指标内容项
    {

        private 焦炭指标内容项数据 currentData;
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
        public double? C
        {
            get
            {
                return this.currentData.C;
            }
            set
            {
                this.currentData.C = value;
            }
        }
        [Bindable(true)]
        public double? HUIFEN
        {
            get
            {
                return this.currentData.HUIFEN;
            }
            set
            {
                this.currentData.HUIFEN = value;
            }
        }
        [Bindable(true)]
        public double? HUIFA
        {
            get
            {
                return this.currentData.HUIFA;
            }
            set
            {
                this.currentData.HUIFA = value;
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
        public double? M40
        {
            get
            {
                return this.currentData.M40;
            }
            set
            {
                this.currentData.M40 = value;
            }
        }
        [Bindable(true)]
        public double? M10
        {
            get
            {
                return this.currentData.M10;
            }
            set
            {
                this.currentData.M10 = value;
            }
        }
        [Bindable(true)]
        public double? QD
        {
            get
            {
                return this.currentData.QD;
            }
            set
            {
                this.currentData.QD = value;
            }
        }
    }
}
