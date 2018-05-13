using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.Data
{

    /// <summary>
    /// 出铁计划
    /// </summary>
    public class ChuTieJiHua : ObservableObject
    {
        private int _gaoLu;
        /// <summary>
        /// 高炉
        /// </summary>
        public int GaoLu
        {
            get { return _gaoLu; }
            set
            {
                if (_gaoLu != value)
                {
                    _gaoLu = value;
                    RaisePropertyChanged("GaoLu");
                }
            }
        }

        private int _xuHao;
        /// <summary>
        /// 序号
        /// </summary>
        public int XuHao
        {
            get { return _xuHao; }
            set
            {
                if (_xuHao != value)
                {
                    _xuHao = value;
                    RaisePropertyChanged("XuHao");
                }
            }
        }

        private int _zdsj;
        /// <summary>
        /// 正点时间
        /// </summary>
        public int Zdsj
        {
            get { return _zdsj; }
            set
            {
                if (_zdsj != value)
                {
                    _zdsj = value;
                    RaisePropertyChanged("Zdsj");
                }
            }
        }

        private int _offSet;
        /// <summary>
        /// 偏移量
        /// </summary>
        public int OffSet
        {
            get { return _offSet; }
            set
            {
                if (_offSet != value)
                {
                    _offSet = value;
                    RaisePropertyChanged("OffSet");
                }
            }
        }


        private string _banCi;
        /// <summary>
        /// 班次
        /// </summary>
        public string BanCi
        {
            get { return _banCi; }
            set
            {
                if (_banCi != value)
                {
                    _banCi = value;
                    RaisePropertyChanged("BanCi");
                }
            }
        }


        private int _banXuHao;
        /// <summary>
        /// 班序号
        /// </summary>
        public int BanXuHao
        {
            get { return _banXuHao; }
            set
            {
                if (_banXuHao != value)
                {
                    _banXuHao = value;
                    RaisePropertyChanged("BanXuHao");
                }
            }
        }

    }
}
