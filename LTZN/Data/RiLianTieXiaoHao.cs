using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.Data
{
    public class RiLianTieXiaoHao : ObservableObject
    {
        private DateTime _riQi;
        /// <summary>
        /// 日期
        /// </summary>
        public DateTime RiQi
        {
            get { return _riQi; }
            set
            {
                if (_riQi != value)
                {
                    _riQi = value;
                    RaisePropertyChanged("RiQi");
                }
            }
        }


        private double? _ziChanJiaoShuiFen;
        /// <summary>
        /// 自产焦水份
        /// </summary>
        public double? ZiChanJiaoShuiFen
        {
            get { return _ziChanJiaoShuiFen; }
            set
            {
                if (_ziChanJiaoShuiFen != value)
                {
                    _ziChanJiaoShuiFen = value;
                    RaisePropertyChanged("ZiChanJiaoShuiFen");
                }
            }
        }

        private double? _luoDiJiaoShuiFen;
        /// <summary>
        /// 落地焦水份
        /// </summary>
        public double? LuoDiJiaoShuiFen
        {
            get { return _luoDiJiaoShuiFen; }
            set
            {
                if (_luoDiJiaoShuiFen != value)
                {
                    _luoDiJiaoShuiFen = value;
                    RaisePropertyChanged("LuoDiJiaoShuiFen");
                }
            }
        }


        private double? _jiaoFenShuiFen;
        /// <summary>
        /// 焦粉水份
        /// </summary>
        public double? JiaoFenShuiFen
        {
            get { return _jiaoFenShuiFen; }
            set
            {
                if (_jiaoFenShuiFen != value)
                {
                    _jiaoFenShuiFen = value;
                    RaisePropertyChanged("JiaoFenShuiFen");
                }
            }
        }


        private double? _erHaoPiDai;
        /// <summary>
        /// 二号皮带
        /// </summary>
        public double? ErHaoPiDai
        {
            get { return _erHaoPiDai; }
            set
            {
                if (_erHaoPiDai != value)
                {
                    _erHaoPiDai = value;
                    RaisePropertyChanged("ErHaoPiDai");
                }
            }
        }


        private double? _sanHaoPiDai;
        /// <summary>
        /// 三号皮带
        /// </summary>
        public double? SanHaoPiDai
        {
            get { return _sanHaoPiDai; }
            set
            {
                if (_sanHaoPiDai != value)
                {
                    _sanHaoPiDai = value;
                    RaisePropertyChanged("SanHaoPiDai");
                }
            }
        }

        private double? _zongFanKuang;
        /// <summary>
        /// 总反矿
        /// </summary>
        public double? ZongFanKuang
        {
            get { return _zongFanKuang; }
            set
            {
                if (_zongFanKuang != value)
                {
                    _zongFanKuang = value;
                    RaisePropertyChanged("ZongFanKuang");
                }
            }
        }


        private double? _shaojieErHaoCheng;
        /// <summary>
        /// 烧结二号称
        /// </summary>
        public double? ShaojieErHaoCheng
        {
            get { return _shaojieErHaoCheng; }
            set
            {
                if (_shaojieErHaoCheng != value)
                {
                    _shaojieErHaoCheng = value;
                    RaisePropertyChanged("ShaojieErHaoCheng");
                }
            }
        }


        private string _beiZhu1;
        /// <summary>
        /// 备注1
        /// </summary>
        public string BeiZhu1
        {
            get { return _beiZhu1; }
            set
            {
                if (_beiZhu1 != value)
                {
                    _beiZhu1 = value;
                    RaisePropertyChanged("BeiZhu1");
                }
            }
        }

        private string _beiZhu2;
        /// <summary>
        /// 备注2
        /// </summary>
        public string BeiZhu2
        {
            get { return _beiZhu2; }
            set
            {
                if (_beiZhu2 != value)
                {
                    _beiZhu2 = value;
                    RaisePropertyChanged("BeiZhu2");
                }
            }
        }

    }
}
