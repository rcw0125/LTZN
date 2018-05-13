using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.Data
{
    public class ChuTie : ObservableObject
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

        private DateTime _zdsj = DateTime.Now;
        /// <summary>
        /// 正点时间
        /// </summary>
        public DateTime Zdsj
        {
            get { return _zdsj; }
            set
            {
                if (_zdsj != value)
                {
                    _zdsj = value;
                    RaisePropertyChanged("Zdsj");
                    CalWdsj();
                }
            }
        }

        private string _luCiHao;
        /// <summary>
        /// 炉次号
        /// </summary>
        public string LuCiHao
        {
            get { return _luCiHao; }
            set
            {
                if (_luCiHao != value)
                {
                    _luCiHao = value;
                    RaisePropertyChanged("LuCiHao");
                }
            }
        }


        private string _banCi="";
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
        /// 本班炉次序号
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


        private DateTime? _dgsj;
        /// <summary>
        /// 对罐时间
        /// </summary>
        public DateTime? Dgsj
        {
            get { return _dgsj; }
            set
            {
                DateTime? val = XiuZheng(value);
                if (_dgsj != val)
                {
                    _dgsj = val;
                    RaisePropertyChanged("Dgsj");
                    CalWdsj();
                }
            }
        }


        private DateTime? _dksj;
        /// <summary>
        /// 堵口时间
        /// </summary>
        public DateTime? Dksj
        {
            get { return _dksj; }
            set
            {
                DateTime? val = XiuZheng(value);
                if (_dksj != val)
                {
                    _dksj = val;
                    RaisePropertyChanged("Dksj");
                    CalWdsj();
                }
            }
        }

       private DateTime? _tzsj;
        /// <summary>
        /// 通知时间
        /// </summary>
        public DateTime? Tzsj
        {
            get { return _tzsj; }
            set
            {
                DateTime? val = XiuZheng(value);
                if (_tzsj != val)
                {
                    _tzsj = val;
                    RaisePropertyChanged("Kksj");
                }
            }
        }

        private int? _wdsj;
        /// <summary>
        ///  晚点时间(分钟)
        /// </summary>
        public int? Wdsj
        {
            get { return _wdsj; }
            set
            {
                if (_wdsj != value)
                {
                    _wdsj = value;
                    RaisePropertyChanged("Wdsj");
                }
            }
        }

        private double? _chanLiang;
        /// <summary>
        ///  产量
        /// </summary>
        public double? ChanLiang
        {
            get { return _chanLiang; }
            set
            {
                if (_chanLiang != value)
                {
                    _chanLiang = value;
                    RaisePropertyChanged("ChanLiang");
                }
            }
        }

        private string _quchu = "";
        /// <summary>
        /// 去处
        /// </summary>
        public string QuChu
        {
            get { return _quchu; }
            set
            {
                if (_quchu != value)
                {
                    _quchu = value;
                    RaisePropertyChanged("QuChu");
                }
            }
        }

        private double? _feC;
        /// <summary>
        /// 铁成分-C
        /// </summary>
        public double? FeC
        {
            get { return _feC; }
            set
            {
                if (_feC != value)
                {
                    _feC = value;
                    RaisePropertyChanged("FeC");
                }
            }
        }

        private double? _feSi;
        /// <summary>
        /// 铁成分-Si
        /// </summary>
        public double? FeSi
        {
            get { return _feSi; }
            set
            {
                if (_feSi != value)
                {
                    _feSi = value;
                    RaisePropertyChanged("FeSi");
                }
            }
        }

        private double? _feMn;
        /// <summary>
        /// 铁成分-Mn
        /// </summary>
        public double? FeMn
        {
            get { return _feMn; }
            set
            {
                if (_feMn != value)
                {
                    _feMn = value;
                    RaisePropertyChanged("FeMn");
                }
            }
        }

        private double? _feP;
        /// <summary>
        ///  铁成分-P
        /// </summary>
        public double? FeP
        {
            get { return _feP; }
            set
            {
                if (_feP != value)
                {
                    _feP = value;
                    RaisePropertyChanged("FeP");
                }
            }
        }


        private double? _feS;
        /// <summary>
        /// 铁成分-S
        /// </summary>
        public double? FeS
        {
            get { return _feS; }
            set
            {
                if (_feS != value)
                {
                    _feS = value;
                    RaisePropertyChanged("FeS");
                }
            }
        }

        private double? _feTi;
        /// <summary>
        ///  铁成分-Ti
        /// </summary>
        public double? FeTi
        {
            get { return _feTi; }
            set
            {
                if (_feTi != value)
                {
                    _feTi = value;
                    RaisePropertyChanged("FeTi");
                }
            }
        }


        private double? _zhaSiO2;
        /// <summary>
        /// 渣-二氧化硅
        /// </summary>
        public double? ZhaSiO2
        {
            get { return _zhaSiO2; }
            set
            {
                if (_zhaSiO2 != value)
                {
                    _zhaSiO2 = value;
                    RaisePropertyChanged("ZhaSiO2");
                }
            }
        }
        

        private double? _zhaCaO;
        /// <summary>
        /// 渣-氧化钙
        /// </summary>
        public double? ZhaCaO
        {
            get { return _zhaCaO; }
            set
            {
                if (_zhaCaO != value)
                {
                    _zhaCaO = value;
                    RaisePropertyChanged("ZhaCaO");
                }
            }
        }

        private double? _zhaMgO;
        /// <summary>
        /// 渣-氧化镁
        /// </summary>
        public double? ZhaMgO
        {
            get { return _zhaMgO; }
            set
            {
                if (_zhaMgO != value)
                {
                    _zhaMgO = value;
                    RaisePropertyChanged("ZhaMgO");
                }
            }
        }

        private double? _zhaAl2O3;
        /// <summary>
        /// 渣-三氧化二铝
        /// </summary>
        public double? ZhaAl2O3
        {
            get { return _zhaAl2O3; }
            set
            {
                if (_zhaAl2O3 != value)
                {
                    _zhaAl2O3 = value;
                    RaisePropertyChanged("ZhaAl2O3");
                }
            }
        }

        private double? _zhaS;
        /// <summary>
        /// 渣-硫
        /// </summary>
        public double? ZhaS
        {
            get { return _zhaS; }
            set
            {
                if (_zhaS != value)
                {
                    _zhaS = value;
                    RaisePropertyChanged("ZhaS");
                }
            }
        }

        private double? _zhaTiO2;
        /// <summary>
        /// 渣-二氧化钛
        /// </summary>
        public double? ZhaTiO2
        {
            get { return _zhaTiO2; }
            set
            {
                if (_zhaTiO2 != value)
                {
                    _zhaTiO2 = value;
                    RaisePropertyChanged("ZhaTiO2");
                }
            }
        }

        private double? _zhaR2;
        /// <summary>
        /// 渣-碱度
        /// </summary>
        public double? ZhaR2
        {
            get { return _zhaR2; }
            set
            {
                if (_zhaR2 != value)
                {
                    _zhaR2 = value;
                    RaisePropertyChanged("ZhaR2");
                }
            }
        }

        private double? _zhaR3;
        /// <summary>
        /// 渣-3元碱度
        /// </summary>
        public double? ZhaR3
        {
            get { return _zhaR3; }
            set
            {
                if (_zhaR3 != value)
                {
                    _zhaR3 = value;
                    RaisePropertyChanged("ZhaR3");
                }
            }
        }

        private double? _zhaR4;
        /// <summary>
        /// 渣-4元碱度
        /// </summary>
        public double? ZhaR4
        {
            get { return _zhaR4; }
            set
            {
                if (_zhaR4 != value)
                {
                    _zhaR4 = value;
                    RaisePropertyChanged("ZhaR4");
                }
            }
        }


        private double? _zhaMgOAlO;

        public double? ZhaMgOAlO
        {
            get { return _zhaMgOAlO; }
            set
            {
                if (_zhaMgOAlO != value)
                {
                    _zhaMgOAlO = value;
                    RaisePropertyChanged("ZhaMgOAlO");
                }
            }
        }

        private int? _liaoPiShu;
        /// <summary>
        /// 料批数
        /// </summary>
        public int? LiaoPiShu
        {
            get { return _liaoPiShu; }
            set
            {
                if (_liaoPiShu != value)
                {
                    _liaoPiShu = value;
                    RaisePropertyChanged("LiaoPiShu");
                }
            }
        }

        private DateTime? XiuZheng(DateTime? time)
        {
            if (time.HasValue)
            {
                DateTime result = new DateTime(Zdsj.Year, Zdsj.Month, Zdsj.Day, time.Value.Hour, time.Value.Minute, 0);
                if (result > Zdsj)
                {
                    TimeSpan span = result - Zdsj;
                    if (span.TotalHours > 16)
                        result = result.AddHours(-1);

                }
                else
                {
                    TimeSpan span = Zdsj - result;
                    if (span.TotalHours > 16)
                        result = result.AddHours(1);
                }
                return result;
            }
            else
                return null;
        }

        private void CalWdsj()
        {
            if (Dgsj != null && Dksj != null)
            {
                if (Dksj.Value < Zdsj)
                    Wdsj = 0;
                else
                {
                    int kaohe = 50;
                    if (GaoLu > 3)
                        kaohe = 60;
                    TimeSpan jiange = Dksj.Value - Dgsj.Value;
                    if (jiange.TotalMinutes > kaohe)
                        Wdsj = Convert.ToInt32(jiange.TotalMinutes) - kaohe;
                    else
                        Wdsj = 0;
                }
            }
            else
                Wdsj = null;
        }
   
    }
}
