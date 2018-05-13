using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.高炉燃料比综合分析
{
    class DataZong
    {
        //序号
        private Int32 num;

        public Int32 Num
        {
            get { return num; }
            set { num = value; }
        }
        // 日期
        private DateTime rq;

        public DateTime Rq
        {
            get { return rq; }
            set { rq = value; }
        }
        //自产焦炭灰分
        private double zichanjthf;

        public double Zichanjthf
        {
            get { return zichanjthf; }
            set { zichanjthf = value; }
        }
        //自产焦炭水分
        private double zichanjtsf;

        public double Zichanjtsf
        {
            get { return zichanjtsf; }
            set { zichanjtsf = value; }
        }
        //外购焦炭灰分
        private double waigoujthf;

        public double Waigoujthf
        {
            get { return waigoujthf; }
            set { waigoujthf = value; }
        }
        //外购焦炭水分
        private double waigoujtsf;

        public double Waigoujtsf
        {
            get { return waigoujtsf; }
            set { waigoujtsf = value; }
        }
        //喷吹煤粉灰分
        private double pcmfhf;

        public double Pcmfhf
        {
            get { return pcmfhf; }
            set { pcmfhf = value; }
        }
        //喷吹煤粉水分
        private double pcmfsf;

        public double Pcmfsf
        {
            get { return pcmfsf; }
            set { pcmfsf = value; }
        }
        //喷吹煤粉挥发分
        private double pcmfhff;

        public double Pcmfhff
        {
            get { return pcmfhff; }
            set { pcmfhff = value; }
        }
        //烧结矿品味
        private double sjkpw;

        public double Sjkpw
        {
            get { return sjkpw; }
            set { sjkpw = value; }
        }
        //球团矿品味
        private double qtkpw;

        public double Qtkpw
        {
            get { return qtkpw; }
            set { qtkpw = value; }
        }
        //澳矿品味
        private double akpw;

        public double Akpw
        {
            get { return akpw; }
            set { akpw = value; }
        }
        //其它块矿品味
        private double qtkkpw;

        public double Qtkkpw
        {
            get { return qtkkpw; }
            set { qtkkpw = value; }
        }
        //钛球品味
        private double tqpw;

        public double Tqpw
        {
            get { return tqpw; }
            set { tqpw = value; }
        }
        //烧结矿配比
        private double sjkpb;

        public double Sjkpb
        {
            get { return sjkpb; }
            set { sjkpb = value; }
        }
        //球团矿配比
        private double qhkpb;
        
        public double Qhkpb
        {
            get { return qhkpb; }
            set { qhkpb = value; }
        }
        //澳矿配比
        private double akpb;

        public double Akpb
        {
            get { return akpb; }
            set { akpb = value; }
        }
        //其它块矿配比
        private double qtkkpb;

        public double Qtkkpb
        {
            get { return qtkkpb; }
            set { qtkkpb = value; }
        }
        //钛球配比
        private double tqpb;

        public double Tqpb
        {
            get { return tqpb; }
            set { tqpb = value; }
        }
        //综合品味
        private double zhpw;

        public double Zhpw
        {
            get { return zhpw; }
            set { zhpw = value; }
        }
        //熔剂配比
        private double rjpb;

        public double Rjpb
        {
            get { return rjpb; }
            set { rjpb = value; }
        }
        //热风温度
        private double rfwd;

        public double Rfwd
        {
            get { return rfwd; }
            set { rfwd = value; }
        }
        //平均顶压
        private double pjdy;

        public double Pjdy
        {
            get { return pjdy; }
            set { pjdy = value; }
        }
        //富氧率
        private double fyl;

        public double Fyl
        {
            get { return fyl; }
            set { fyl = value; }
        }
        //生铁含si
        private double fesi;

        public double Fesi
        {
            get { return fesi; }
            set { fesi = value; }
        }
        //S负荷
        private double sfh;

        public double Sfh
        {
            get { return sfh; }
            set { sfh = value; }
        }
        //炉渣碱度
        private double lzjd;

        public double Lzjd
        {
            get { return lzjd; }
            set { lzjd = value; }
        }
        //日产品
        private double rcl;

        public double Rcl
        {
            get { return rcl; }
            set { rcl = value; }
        }
        //理论回收率
        private double llhsl;

        public double Llhsl
        {
            get { return llhsl; }
            set { llhsl = value; }
        }
        //实际回收率
        private double sjhsl;

        public double Sjhsl
        {
            get { return sjhsl; }
            set { sjhsl = value; }
        }
        //焦比
        private double jb;

        public double Jb
        {
            get { return jb; }
            set { jb = value; }
        }
        //燃料比
        private double rlb;

        public double Rlb
        {
            get { return rlb; }
            set { rlb = value; }
        }
        //煤比
        private double mb;

        public double Mb
        {
            get { return mb; }
            set { mb = value; }
        }
      //含铁原料
        private double htyl;

        public double Htyl
        {
            get { return htyl; }
            set { htyl = value; }
        }
        
        
        //理论铁量
        private double lltl;

        public double Lltl
        {
            get { return lltl; }
            set { lltl = value; }
        
        
        }
    }
}
