using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;

namespace LTZN.Data
{
    /// <summary>
    /// 月能耗
    /// </summary>
    public class MonthEnergy : ObservableObject
    {
        private int _nian;
        /// <summary>
        /// 年
        /// </summary>
        public int Nian
        {
            get { return _nian; }
            set
            {
                if (_nian != value)
                {
                    _nian = value;
                    RaisePropertyChanged("Nian");
                }
            }
        }

        private int _yue;
        /// <summary>
        /// 月
        /// </summary>
        public int Yue
        {
            get { return _yue; }
            set
            {
                if (_yue != value)
                {
                    _yue = value;
                    RaisePropertyChanged("Yue");
                }
            }
        }

        private double? _ganMaoJiao;
        /// <summary>
        /// 干毛焦
        /// </summary>
        public double? GanMaoJiao
        {
            get { return _ganMaoJiao; }
            set
            {
                if (_ganMaoJiao != value)
                {
                    _ganMaoJiao = value;
                    RaisePropertyChanged("GanMaoJiao");
                }
            }
        }

        private double? _shiJiaoFen;
        /// <summary>
        ///  湿焦粉
        /// </summary>
        public double? ShiJiaoFen
        {
            get { return _shiJiaoFen; }
            set
            {
                if (_shiJiaoFen != value)
                {
                    _shiJiaoFen = value;
                    RaisePropertyChanged("ShuiJiaoFen");
                }
            }
        }

        private double? _jiaoDing;
        /// <summary>
        /// 焦丁
        /// </summary>
        public double? JiaoDing
        {
            get { return _jiaoDing; }
            set
            {
                if (_jiaoDing != value)
                {
                    _jiaoDing = value;
                    RaisePropertyChanged("JiaoDing");
                }
            }
        }


        private double? _shiMeiFen;
        /// <summary>
        /// 湿煤粉
        /// </summary>
        public double? ShiMeiFen
        {
            get { return _shiMeiFen; }
            set
            {
                if (_shiMeiFen != value)
                {
                    _shiMeiFen = value;
                    RaisePropertyChanged("ShuiMeiFen");
                }
            }
        }


        private double? _haoDianLiang;
        /// <summary>
        /// 耗电量
        /// </summary>
        public double? HaoDianLiang
        {
            get { return _haoDianLiang; }
            set
            {
                if (_haoDianLiang != value)
                {
                    _haoDianLiang = value;
                    RaisePropertyChanged("HaoDianLiang");
                }
            }
        }

        private double? _zhengQi;
        /// <summary>
        /// 蒸汽
        /// </summary>
        public double? ZhengQi
        {
            get { return _zhengQi; }
            set
            {
                if (_zhengQi != value)
                {
                    _zhengQi = value;
                    RaisePropertyChanged("ZhengQi");
                }
            }
        }

        private double? _haoShui;
        /// <summary>
        /// 耗水
        /// </summary>
        public double? HaoShui
        {
            get { return _haoShui; }
            set
            {
                if (_haoShui != value)
                {
                    _haoShui = value;
                    RaisePropertyChanged("HaoShui");
                }
            }
        }

        private double? _gaoLuMeiQi;
        /// <summary>
        /// 高炉煤气
        /// </summary>
        public double? GaoLuMeiQi
        {
            get { return _gaoLuMeiQi; }
            set
            {
                if (_gaoLuMeiQi != value)
                {
                    _gaoLuMeiQi = value;
                    RaisePropertyChanged("GaoLuMeiQi");
                }
            }
        }

        private double? _huiShouMeiQi;
        /// <summary>
        /// 回收煤气
        /// </summary>
        public double? HuiShouMeiQi
        {
            get { return _huiShouMeiQi; }
            set
            {
                if (_huiShouMeiQi != value)
                {
                    _huiShouMeiQi = value;
                    RaisePropertyChanged("HuiShouMeiQi");
                }
            }
        }

        private double? _fuYang;
        /// <summary>
        /// 富氧
        /// </summary>
        public double? FuYang
        {
            get { return _fuYang; }
            set
            {
                if (_fuYang != value)
                {
                    _fuYang = value;
                    RaisePropertyChanged("FuYang");
                }
            }
        }

        private double? _danQi;
        /// <summary>
        /// 氮气
        /// </summary>
        public double? DanQi
        {
            get { return _danQi; }
            set
            {
                if (_danQi != value)
                {
                    _danQi = value;
                    RaisePropertyChanged("DanQi");
                }
            }
        }


        private double? _daKuaiTie;
        /// <summary>
        /// 大块铁
        /// </summary>
        public double? DaKuaiTie
        {
            get { return _daKuaiTie; }
            set
            {
                if (_daKuaiTie != value)
                {
                    _daKuaiTie = value;
                    RaisePropertyChanged("DaKuaiTie");
                }
            }
        }

        private int? _gaoLu1RenShu;
        /// <summary>
        /// 高炉1人数
        /// </summary>
        public int? GaoLu1RenShu
        {
            get { return _gaoLu1RenShu; }
            set
            {
                if (_gaoLu1RenShu != value)
                {
                    _gaoLu1RenShu = value;
                    RaisePropertyChanged("GaoLu1RenShu");
                }
            }
        }

        private int? _gaoLu2RenShu;
        /// <summary>
        /// 高炉2人数
        /// </summary>
        public int? GaoLu2RenShu
        {
            get { return _gaoLu2RenShu; }
            set
            {
                if (_gaoLu2RenShu != value)
                {
                    _gaoLu2RenShu = value;
                    RaisePropertyChanged("GaoLu2RenShu");
                }
            }
        }

        private int? _gaoLu3RenShu;
        /// <summary>
        /// 高炉3人数
        /// </summary>
        public int? GaoLu3RenShu
        {
            get { return _gaoLu3RenShu; }
            set
            {
                if (_gaoLu3RenShu != value)
                {
                    _gaoLu3RenShu = value;
                    RaisePropertyChanged("GaoLu3RenShu");
                }
            }
        }

        private int? _gaoLu4RenShu;
        /// <summary>
        /// 高炉4人数
        /// </summary>
        public int? GaoLu4RenShu
        {
            get { return _gaoLu4RenShu; }
            set
            {
                if (_gaoLu4RenShu != value)
                {
                    _gaoLu4RenShu = value;
                    RaisePropertyChanged("GaoLu4RenShu");
                }
            }
        }

        private int? _gaoLu5RenShu;
        /// <summary>
        /// 高炉5人数
        /// </summary>
        public int? GaoLu5RenShu
        {
            get { return _gaoLu5RenShu; }
            set
            {
                if (_gaoLu5RenShu != value)
                {
                    _gaoLu5RenShu = value;
                    RaisePropertyChanged("GaoLu5RenShu");
                }
            }
        }

        private int? _gaoLu6RenShu;
        /// <summary>
        /// 高炉6人数
        /// </summary>
        public int? GaoLu6RenShu
        {
            get { return _gaoLu6RenShu; }
            set
            {
                if (_gaoLu6RenShu != value)
                {
                    _gaoLu6RenShu = value;
                    RaisePropertyChanged("GaoLu6RenShu");
                }
            }
        }


        #region 焦成分

        private double? _jiaoShuiFen;
        /// <summary>
        /// 焦水份
        /// </summary>
        public double? JiaoShuiFen
        {
            get { return _jiaoShuiFen; }
            set
            {
                if (_jiaoShuiFen != value)
                {
                    _jiaoShuiFen = value;
                    RaisePropertyChanged("JiaoShuiFen");
                }
            }
        }

        private double? _jiaoHuiFen;
        /// <summary>
        /// 焦灰份
        /// </summary>
        public double? JiaoHuiFen
        {
            get { return _jiaoHuiFen; }
            set
            {
                if (_jiaoHuiFen != value)
                {
                    _jiaoHuiFen = value;
                    RaisePropertyChanged("JiaoHuiFen");
                }
            }
        }

        private double? _jiaoHuiFaFen;
        /// <summary>
        /// 挥发份
        /// </summary>
        public double? JiaoHuiFaFen
        {
            get { return _jiaoHuiFaFen; }
            set
            {
                if (_jiaoHuiFaFen != value)
                {
                    _jiaoHuiFaFen = value;
                    RaisePropertyChanged("JiaoHuiFaFen");
                }
            }
        }

        private double? _jiaoS;
        /// <summary>
        /// 焦S
        /// </summary>
        public double? JiaoS
        {
            get { return _jiaoS; }
            set
            {
                if (_jiaoS != value)
                {
                    _jiaoS = value;
                    RaisePropertyChanged("JiaoS");
                }
            }
        }

        private double? _jiaoC;
        /// <summary>
        /// 焦C
        /// </summary>
        public double? JiaoC
        {
            get { return _jiaoC; }
            set
            {
                if (_jiaoC != value)
                {
                    _jiaoC = value;
                    RaisePropertyChanged("JiaoC");
                }
            }
        }

        private double? _jiaoM25;
        /// <summary>
        /// 焦M25
        /// </summary>
        public double? JiaoM25
        {
            get { return _jiaoM25; }
            set
            {
                if (_jiaoM25 != value)
                {
                    _jiaoM25 = value;
                    RaisePropertyChanged("JiaoM25");
                }
            }
        }

        private double? _jiaoM10;
        /// <summary>
        /// 焦M10
        /// </summary>
        public double? JiaoM10
        {
            get { return _jiaoM10; }
            set
            {
                if (_jiaoM10 != value)
                {
                    _jiaoM10 = value;
                    RaisePropertyChanged("JiaoM10");
                }
            }
        }

        #endregion


        #region 煤成分
        private double? _meiShuiFen;
        /// <summary>
        /// 煤水份
        /// </summary>
        public double? MeiShuiFen
        {
            get { return _meiShuiFen; }
            set
            {
                if (_meiShuiFen != value)
                {
                    _meiShuiFen = value;
                    RaisePropertyChanged("MeiShuiFen");
                }
            }
        }

        private double? _meiHuiFen;
        /// <summary>
        /// 煤灰份
        /// </summary>
        public double? MeiHuiFen
        {
            get { return _meiHuiFen; }
            set
            {
                if (_meiHuiFen != value)
                {
                    _meiHuiFen = value;
                    RaisePropertyChanged("MeiHuiFen");
                }
            }
        }

        private double? _meiHuiFaFen;
        /// <summary>
        /// 煤挥发份
        /// </summary>
        public double? MeiHuiFaFen
        {
            get { return _meiHuiFaFen; }
            set
            {
                if (_meiHuiFaFen != value)
                {
                    _meiHuiFaFen = value;
                    RaisePropertyChanged("MeiHuiFaFen");
                }
            }
        }

        private double? _meiS;
        /// <summary>
        /// 煤S
        /// </summary>
        public double? MeiS
        {
            get { return _meiS; }
            set
            {
                if (_meiS != value)
                {
                    _meiS = value;
                    RaisePropertyChanged("MeiS");
                }
            }
        }

        #endregion
    }
}
