using System;
using System.Collections.Generic;
using System.Text;
using LTZN.Data;
using System.Data;
using System.Data.OracleClient;

namespace LTZN.调度
{
    class YLXiaoHao : DataObject
    {
        private decimal? _ShaoJieKuang = null;
        /// <summary>
        /// 烧结矿
        /// </summary>
        public decimal? ShaoJieKuang
        {
            get
            {
                return _ShaoJieKuang;
            }
            set
            {
                if (_ShaoJieKuang != value)
                {
                    _ShaoJieKuang = value;
                    RaisePropertyChanged("ShaoJieKuang");
                }
            }
        }

        private decimal? _ShuLuQiu = null;
        /// <summary>
        /// 竖炉球
        /// </summary>
        public decimal? ShuLuQiu
        {
            get
            {
                return _ShuLuQiu;
            }
            set
            {
                if (_ShuLuQiu != value)
                {
                    _ShuLuQiu = value;
                    RaisePropertyChanged("ShuLuQiu");
                }
            }
        }


        private decimal? _BenXiKuang = null;
        /// <summary>
        /// 本溪矿
        /// </summary>
        public decimal? BenXiKuang
        {
            get
            {
                return _BenXiKuang;
            }
            set
            {
                if (_BenXiKuang != value)
                {
                    _BenXiKuang = value;
                    RaisePropertyChanged("BenXiKuang");
                }
            }
        }


      

        private decimal? _GongYiCheng = null;
        /// <summary>
        /// 工艺称
        /// </summary>
        public decimal? GongYiCheng
        {
            get
            {
                return _GongYiCheng;
            }
            set
            {
                if (_GongYiCheng != value)
                {
                    _GongYiCheng = value;
                    RaisePropertyChanged("GongYiCheng");
                }
            }
        }


        private decimal? _JiaoDing = null;
        /// <summary>
        /// 焦丁
        /// </summary>
        public decimal? JiaoDing
        {
            get
            {
                return _JiaoDing;
            }
            set
            {
                if (_JiaoDing != value)
                {
                    _JiaoDing = value;
                    RaisePropertyChanged("JiaoDing");
                }
            }
        }

        private decimal? _PenMei = null;
        /// <summary>
        /// 喷煤 煤粉
        /// </summary>
        public decimal? PenMei
        {
            get
            {
                return _PenMei;
            }
            set
            {
                if (_PenMei != value)
                {
                    _PenMei = value;
                    RaisePropertyChanged("PenMei");
                }
            }
        }

        private decimal? _FuYangLiang = null;
        /// <summary>
        /// 富氧量
        /// </summary>
        public decimal? FuYangLiang
        {
            get
            {
                return _FuYangLiang;
            }
            set
            {
                if (_FuYangLiang != value)
                {
                    _FuYangLiang = value;
                    RaisePropertyChanged("FuYangLiang");
                }
            }
        }

        private decimal? _PBKuai = null;
        /// <summary>
        /// PB块
        /// </summary>
        public decimal? PBKuai
        {
            get
            {
                return _PBKuai;
            }
            set
            {
                if (_PBKuai != value)
                {
                    _PBKuai = value;
                    RaisePropertyChanged("PBKuai");
                }
            }
        }

        private decimal? _FMGKuai = null;
        /// <summary>
        /// FMG块
        /// </summary>
        public decimal? FMGKuai
        {
            get
            {
                return _FMGKuai;
            }
            set
            {
                if (_FMGKuai != value)
                {
                    _FMGKuai = value;
                    RaisePropertyChanged("FMGKuai");
                }
            }
        }

        private decimal? _GuiShi = null;
        /// <summary>
        /// 硅石
        /// </summary>
        public decimal? GuiShi
        {
            get
            {
                return _GuiShi;
            }
            set
            {
                if (_GuiShi != value)
                {
                    _GuiShi = value;
                    RaisePropertyChanged("GuiShi");
                }
            }
        }

        private decimal? _SheWenShi = null;
        /// <summary>
        /// 蛇纹石
        /// </summary>
        public decimal? SheWenShi
        {
            get
            {
                return _SheWenShi;
            }
            set
            {
                if (_SheWenShi != value)
                {
                    _SheWenShi = value;
                    RaisePropertyChanged("SheWenShi");
                }
            }
        }

        private decimal? _YingShi = null;
        /// <summary>
        /// 萤石
        /// </summary>
        public decimal? YingShi
        {
            get
            {
                return _YingShi;
            }
            set
            {
                if (_YingShi != value)
                {
                    _YingShi = value;
                    RaisePropertyChanged("YingShi");
                }
            }
        }

        private decimal? _BaiYunShi = null;
        /// <summary>
        /// 白云石
        /// </summary>
        public decimal? BaiYunShi
        {
            get
            {
                return _BaiYunShi;
            }
            set
            {
                if (_BaiYunShi != value)
                {
                    _BaiYunShi = value;
                    RaisePropertyChanged("BaiYunShi");
                }
            }
        }

        private decimal? _TaiQiu = null;
        /// <summary>
        /// 钛球
        /// </summary>
        public decimal? TaiQiu
        {
            get
            {
                return _TaiQiu;
            }
            set
            {
                if (_TaiQiu != value)
                {
                    _TaiQiu = value;
                    RaisePropertyChanged("TaiQiu");
                }
            }
        }


        private decimal? _MengKuang = null;
        /// <summary>
        /// 锰矿
        /// </summary>
        public decimal? MengKuang
        {
            get
            {
                return _MengKuang;
            }
            set
            {
                if (_MengKuang != value)
                {
                    _MengKuang = value;
                    RaisePropertyChanged("MengKuang");
                }
            }
        }

        private decimal? _QiTaShuLiao = null;
        /// <summary>
        /// 其它熟料
        /// </summary>
        public decimal? QiTaShuLiao
        {
            get
            {
                return _QiTaShuLiao;
            }
            set
            {
                if (_QiTaShuLiao != value)
                {
                    _QiTaShuLiao = value;
                    RaisePropertyChanged("QiTaShuLiao");
                }
            }
        }

        private string _QiTaShuLiaoMC = "";
        /// <summary>
        /// 其它熟料名称
        /// </summary>
        public string QiTaShuLiaoMC
        {
            get
            {
                if (string.IsNullOrEmpty(_QiTaShuLiaoMC))
                    return "其它熟料";
                else
                    return _QiTaShuLiaoMC;
            }
            set
            {
                if (_QiTaShuLiaoMC != value)
                {
                    _QiTaShuLiaoMC = value;
                    RaisePropertyChanged("QiTaShuLiaoMC");
                }
            }
        }

        private decimal? _QiTaShengLiao = null;
        /// <summary>
        /// 其它生料
        /// </summary>
        public decimal? QiTaShengLiao
        {
            get
            {
                    return _QiTaShengLiao;
            }
            set
            {
                if (_QiTaShengLiao != value)
                {
                    _QiTaShengLiao = value;
                    RaisePropertyChanged("QiTaShengLiao");
                }
            }
        }

        private string _QiTaShengLiaoMC = "";
        /// <summary>
        /// 其它生料名称
        /// </summary>
        public string QiTaShengLiaoMC
        {
            get
            {
                if (string.IsNullOrEmpty(_QiTaShengLiaoMC))
                    return "其它生料";
                else
                    return _QiTaShengLiaoMC;
            }
            set
            {
                if (_QiTaShengLiaoMC != value)
                {
                    _QiTaShengLiaoMC = value;
                    RaisePropertyChanged("QiTaShengLiaoMC");
                }
            }
        }

        private decimal? _ZiChanShiJiao = null;
        /// <summary>
        /// 高炉自产湿焦
        /// </summary>
        public decimal? ZiChanShiJiao
        {
            get
            {
                return _ZiChanShiJiao;
            }
            set
            {
                if (_ZiChanShiJiao != value)
                {
                    _ZiChanShiJiao = value;
                    RaisePropertyChanged("ZiChanShiJiao");
                }
            }
        }

        private decimal? _LuoDiShiJiao = null;
        /// <summary>
        /// 高炉落地湿焦
        /// </summary>
        public decimal? LuoDiShiJiao
        {
            get
            {
                return _LuoDiShiJiao;
            }
            set
            {
                if (_LuoDiShiJiao != value)
                {
                    _LuoDiShiJiao = value;
                    RaisePropertyChanged("LuoDiShiJiao");
                }
            }
        }

        private decimal? _QuanChangZiChanShiJiao = null;
        /// <summary>
        /// 全厂自产湿焦
        /// </summary>
        public decimal? QuanChangZiChanShiJiao
        {
            get
            {
                return _QuanChangZiChanShiJiao;
            }
            set
            {
                if (_QuanChangZiChanShiJiao != value)
                {
                    _QuanChangZiChanShiJiao = value;
                    RaisePropertyChanged("QuanChangZiChanShiJiao");
                }
            }
        }

        private decimal? _QuanChangLuoDiShiJiao = null;
        /// <summary>
        /// 全厂高炉落地湿焦
        /// </summary>
        public decimal? QuanChangLuoDiShiJiao
        {
            get
            {
                return _QuanChangLuoDiShiJiao;
            }
            set
            {
                if (_QuanChangLuoDiShiJiao != value)
                {
                    _QuanChangLuoDiShiJiao = value;
                    RaisePropertyChanged("QuanChangLuoDiShiJiao");
                }
            }
        }
        private decimal? _Weikuai = null;
        /// <summary>
        /// 尾块
        /// </summary>
        public decimal? Weikuai
        {
            get
            {
                return _Weikuai;
            }
            set
            {
                if (_Weikuai != value)
                {
                    _Weikuai = value;
                    RaisePropertyChanged("Weikuai");
                }
            }
        
        }
        private decimal? _Malaikuai = null;
        /// <summary>
        /// 马来块
        /// </summary>
        public decimal? Malaikuai
        {
            get
            {
                return _Malaikuai;
            }
            set
            {
                if (_Malaikuai != value)
                {
                    _Malaikuai = value;
                    RaisePropertyChanged("Malaikuai");
                }
            }

        }

        private decimal? _Niumankuai = null;
        /// <summary>
        /// 纽曼块
        /// </summary>
        public decimal? Niumankuai
        {
            get
            {
                return _Niumankuai;
            }
            set
            {
                if (_Niumankuai != value)
                {
                    _Niumankuai = value;
                    RaisePropertyChanged("Niumankuai");
                }
            }

        }

        private decimal? _Zichanqiu = null;
        /// <summary>
        /// 自产球
        /// </summary>
        public decimal? Zichanqiu
        {
            get
            {
                return _Zichanqiu;
            }
            set
            {
                if (_Zichanqiu != value)
                {
                    _Zichanqiu = value;
                    RaisePropertyChanged("Zichanqiu");
                }
            }

        }
        private decimal? _Wantianxinqiu = null;
        /// <summary>
        /// 万天新球
        /// </summary>
        public decimal? Wantianxinqiu
        {
            get
            {
                return _Wantianxinqiu;
            }
            set
            {
                if (_Wantianxinqiu != value)
                {
                    _Wantianxinqiu = value;
                    RaisePropertyChanged("Wantianxinqiu");
                }
            }

        }
        private decimal? _Baxiqiu = null;
        /// <summary>
        /// 巴西球
        /// </summary>
        public decimal? Baxiqiu
        {
            get
            {
                return _Baxiqiu;
            }
            set
            {
                if (_Baxiqiu != value)
                {
                    _Baxiqiu = value;
                    RaisePropertyChanged("Baxiqiu");
                }
            }

        }
        private decimal? _Suitie = null;
        /// <summary>
        /// 高炉碎铁
        /// </summary>
        public decimal? Suitie
        {
            get
            {
                return _Suitie;
            }
            set
            {
                if (_Suitie != value)
                {
                    _Suitie = value;
                    RaisePropertyChanged("Suitie");
                }
            }

        }

         private decimal? _Qiutuankuang = null;
        /// <summary>
        /// 球团矿
        /// </summary>
        public decimal? Qiutuankuang
        {
            get
            {
                return _Qiutuankuang;
            }
            set
            {
                if (_Qiutuankuang != value)
                {
                    _Qiutuankuang = value;
                    RaisePropertyChanged("Qiutuankuang");
                }
            }

        }

             private decimal? _Gaotaiqiutuankuang = null;
        /// <summary>
        /// 高钛球团矿
        /// </summary>
        public decimal? Gaotaiqiutuankuang
        {
            get
            {
                return _Gaotaiqiutuankuang;
            }
            set
            {
                if (_Gaotaiqiutuankuang != value)
                {
                    _Gaotaiqiutuankuang = value;
                    RaisePropertyChanged("Gaotaiqiutuankuang");
                }
            }

        }

                     private decimal? _Guoneiqiutuankuang = null;
        /// <summary>
        /// 国内球团矿
        /// </summary>
        public decimal? Guoneiqiutuankuang
        {
            get
            {
                return _Guoneiqiutuankuang;
            }
            set
            {
                if (_Guoneiqiutuankuang != value)
                {
                    _Guoneiqiutuankuang = value;
                    RaisePropertyChanged("Guoneiqiutuankuang");
                }
            }

        }

                             private decimal? _Jinkouqiutuankuang = null;
        /// <summary>
        /// 进口球团矿
        /// </summary>
        public decimal? Jinkouqiutuankuang
        {
            get
            {
                return _Jinkouqiutuankuang;
            }
            set
            {
                if (_Jinkouqiutuankuang != value)
                {
                    _Jinkouqiutuankuang = value;
                    RaisePropertyChanged("Jinkouqiutuankuang");
                }
            }

        }

                                    private decimal? _Qitakuaikuang = null;
        /// <summary>
        /// 其它块矿
        /// </summary>
        public decimal? Qitakuaikuang
        {
            get
            {
                return _Qitakuaikuang;
            }
            set
            {
                if (_Qitakuaikuang != value)
                {
                    _Qitakuaikuang = value;
                    RaisePropertyChanged("Qitakuaikuang");
                }
            }

        }

                                           private decimal? _Qitarongji = null;
        /// <summary>
        /// 其它熔剂
        /// </summary>
        public decimal? Qitarongji
        {
            get
            {
                return _Qitarongji;
            }
            set
            {
                if (_Qitarongji != value)
                {
                    _Qitarongji = value;
                    RaisePropertyChanged("Qitarongji");
                }
            }

        }

        private decimal? _Gaopinweitaiqiu = null;
        /// <summary>
        /// 高品位钛球
        /// </summary>
        public decimal? Gaopinweitaiqiu
        {
            get
            {
                return _Gaopinweitaiqiu;
            }
            set
            {
                if (_Gaopinweitaiqiu != value)
                {
                    _Gaopinweitaiqiu = value;
                    RaisePropertyChanged("Gaopinweitaiqiu");
                }
            }

        }
     
                   
                  
                  
                    
                

        
        /// <summary>
        /// 除了焦炭，其余初始为0，也就是只初始化提取的数据
        /// </summary>
        public void InitNoJiao()
        {
            BaiYunShi = 0;
            BenXiKuang = 0;
            FMGKuai = 0;
            FuYangLiang = 0;
            GuiShi = 0;
            JiaoDing = 0;
            GongYiCheng = 0;
            MengKuang = 0;
            PBKuai = 0;
            PenMei = 0;
            QiTaShengLiao = 0;
            QiTaShengLiaoMC = "";
            QiTaShuLiao = 0;
            QiTaShuLiaoMC = "";
            ShaoJieKuang = 0;
            SheWenShi = 0;
            ShuLuQiu = 0;
            TaiQiu = 0;
            YingShi = 0;
            Weikuai = 0;
            Malaikuai = 0;
            Niumankuai = 0;
            Zichanqiu = 0;
            Wantianxinqiu = 0;
            Baxiqiu = 0;
            Suitie = 0;

            ZiChanShiJiao = 0;
            LuoDiShiJiao = 0;

             Qiutuankuang=0;
             Gaotaiqiutuankuang=0;
             Guoneiqiutuankuang=0;
             Jinkouqiutuankuang=0;
             Qitakuaikuang=0;
             Qitarongji=0;
             Gaopinweitaiqiu=0;


              
        }

        public void ImportData(int gaolu, DateTime Rq)
        {
            InitNoJiao();

            try
            {
                OracleConnection cn = new OracleConnection();
                cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
                cn.Open();
                OracleCommand cmd = new OracleCommand();
                cmd.Connection = cn;
                cmd.CommandText = "SELECT  MC,NVL(BAIBAN,0)+NVL(ZHONGBAN,0)+NVL(YEBAN,0),BEIZHU FROM RBXIAOHAO WHERE TRUNC(SJ)=:RQ AND GAOLU=:GAOLU";
                cmd.Parameters.Add(":RQ", OracleType.DateTime).Value = Rq;
                cmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = gaolu;
                OracleDataReader dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    string mc = "";
                    if (!dr.IsDBNull(0))
                        mc = dr.GetString(0);
                    switch (mc)
                    {
                        case "烧结矿":
                            ShaoJieKuang = dr.GetDecimal(1);
                            break;
                        case "竖炉球":
                            ShuLuQiu = dr.GetDecimal(1);
                            break;
                        case "本溪矿":
                            BenXiKuang = dr.GetDecimal(1);
                            break;
                        case "焦炭":
                            GongYiCheng = dr.GetDecimal(1);
                            break;
                        case "焦丁":
                            JiaoDing = dr.GetDecimal(1);
                            break;
                        case "喷煤":
                            PenMei = dr.GetDecimal(1);
                            break;
                        case "富氧量":
                            FuYangLiang = dr.GetDecimal(1);
                            break;
                        case "PB块":
                            PBKuai = dr.GetDecimal(1);
                            break;
                        case "纽曼块":
                            Niumankuai = dr.GetDecimal(1);
                            break;
                        case "FMG块":
                            FMGKuai = dr.GetDecimal(1);
                            break;
                        case "硅石":
                            GuiShi = dr.GetDecimal(1);
                            break;
                        case "蛇纹石":
                            SheWenShi = dr.GetDecimal(1);
                            break;
                        case "萤石":
                            YingShi = dr.GetDecimal(1);
                            break;
                        case "白云石":
                            BaiYunShi = dr.GetDecimal(1);
                            break;
                        case "钛球":
                            TaiQiu = dr.GetDecimal(1);
                            break;
                        case "锰矿":
                            MengKuang = dr.GetDecimal(1);
                            break;
                        case "其它熟料":
                            QiTaShengLiao = dr.GetDecimal(1);
                            if (!dr.IsDBNull(2))
                                QiTaShengLiaoMC = dr.GetString(2);
                            break;
                        case "其它生料":
                            QiTaShengLiao = dr.GetDecimal(1);
                            if (!dr.IsDBNull(2))
                                QiTaShengLiaoMC = dr.GetString(2);
                            break;
                        case "尾块":
                            Weikuai = dr.GetDecimal(1);
                            break;
                        case "马来块":
                            Malaikuai = dr.GetDecimal(1);
                            break;


                        case "自产球":
                            Zichanqiu = dr.GetDecimal(1);
                            break;
                        case "万天新球":
                            Wantianxinqiu = dr.GetDecimal(1);
                            break;
                        case "巴西球":
                            Baxiqiu = dr.GetDecimal(1);
                            break;
                        case "碎铁":
                            Suitie = dr.GetDecimal(1);
                            break;

                        case "球团矿":
                            Qiutuankuang = dr.GetDecimal(1);
                            break;
                        case "高钛球团矿":
                            Gaotaiqiutuankuang = dr.GetDecimal(1);
                            break;
                        case "国内球团矿":
                            Guoneiqiutuankuang = dr.GetDecimal(1);
                            break;
                        case "进口球团矿":
                            Jinkouqiutuankuang = dr.GetDecimal(1);
                            break;
                        case "其它块矿":
                            Qitakuaikuang = dr.GetDecimal(1);
                            break;
                        case "其它熔剂":
                            Qitarongji = dr.GetDecimal(1);
                            break;
                        case "高品位钛球":
                            Gaopinweitaiqiu = dr.GetDecimal(1);
                            break;



                    }
                }
                dr.Close();
                cn.Close();
            }
            catch
            {

                return;
            
            }

           
        }

        public void LoadData(int gaolu, DateTime Rq)
        {
            InitNoJiao();
            OracleConnection cn = new OracleConnection();
            cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
            cn.Open();
            GaoluLoadData(cn, gaolu, Rq);
          //  RiLoadData(cn, Rq);
            cn.Close();
        }

        public bool GaoluLoadData(OracleConnection cn, int gaolu, DateTime Rq)
        {
            bool find = false;
            OracleCommand selectCmd = new OracleCommand();
            selectCmd.Connection = cn;
            selectCmd.CommandText = "SELECT 机烧矿, 球团矿, 国内球团矿,进口球团矿,PB块,纽曼块,其它块矿,高钛球团矿,高品位钛球,硅石,萤石,蛇纹石,其它熔剂,富氧量,工艺称,煤粉,焦丁,自产湿焦,落地湿焦,生成标志 FROM 原料消耗 WHERE (日期=:RQ AND 高炉=:GAOLU)";
            selectCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = Rq;
            selectCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = gaolu;

            OracleDataReader dr = selectCmd.ExecuteReader();
            if (dr.Read())
            {
                if (dr.IsDBNull(0)) this.ShaoJieKuang = null; else this.ShaoJieKuang = dr.GetDecimal(0);
                if (dr.IsDBNull(1)) this.Qiutuankuang = null; else this.Qiutuankuang = dr.GetDecimal(1);
                if (dr.IsDBNull(2)) this.Guoneiqiutuankuang = null; else this.Guoneiqiutuankuang = dr.GetDecimal(2);
                if (dr.IsDBNull(3)) this.Jinkouqiutuankuang = null; else this.Jinkouqiutuankuang = dr.GetDecimal(3);
                if (dr.IsDBNull(4)) this.PBKuai = null; else this.PBKuai = dr.GetDecimal(4);
                if (dr.IsDBNull(5)) this.Niumankuai = null; else this.Niumankuai = dr.GetDecimal(5);
                if (dr.IsDBNull(6)) this.Qitakuaikuang = null; else Qitakuaikuang = dr.GetDecimal(6);

                if (dr.IsDBNull(7)) this.Gaotaiqiutuankuang = null; else this.Gaotaiqiutuankuang = dr.GetDecimal(7);
                if (dr.IsDBNull(8)) this.Gaopinweitaiqiu = null; else this.Gaopinweitaiqiu = dr.GetDecimal(8);
                if (dr.IsDBNull(9)) this.GuiShi = null; else this.GuiShi = dr.GetDecimal(9);
                if (dr.IsDBNull(10)) this.YingShi = null; else this.YingShi = dr.GetDecimal(10);
                if (dr.IsDBNull(11)) this.SheWenShi = null; else this.SheWenShi = dr.GetDecimal(11);
                if (dr.IsDBNull(12)) this.Qitarongji = null; else this.Qitarongji = dr.GetDecimal(12);
                if (dr.IsDBNull(13)) this.FuYangLiang = null; else this.FuYangLiang = dr.GetDecimal(13);

                if (dr.IsDBNull(14)) this.GongYiCheng = null; else this.GongYiCheng = dr.GetDecimal(14);
                if (dr.IsDBNull(15)) this.PenMei = null; else this.PenMei = dr.GetDecimal(15);
                if (dr.IsDBNull(16)) this.JiaoDing = null; else this.JiaoDing = dr.GetDecimal(16);

                if (dr.IsDBNull(17)) this.ZiChanShiJiao = null; else this.ZiChanShiJiao = dr.GetDecimal(17);
                if (dr.IsDBNull(18)) this.LuoDiShiJiao = null; else this.LuoDiShiJiao = dr.GetDecimal(18);
               
                find = true;
                // if (dr.IsDBNull(15)) item.生成标志 = null; else item.生成标志 = Convert.ToInt32(dr.GetDecimal(15));
            }
            return find;

        }

        public bool RiLoadData(OracleConnection cn,DateTime Rq)
        {
            bool find = false;
            OracleCommand selectCmd = new OracleCommand();
            selectCmd.Connection = cn;
            selectCmd.CommandText = "SELECT 自产湿焦,落地湿焦 FROM 全厂日消耗 WHERE P01日期=:RQ";
            selectCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = Rq;
            OracleDataReader dr = selectCmd.ExecuteReader();
            if (dr.Read())
            {
                if (dr.IsDBNull(0)) this.QuanChangZiChanShiJiao = null; else this.QuanChangZiChanShiJiao = dr.GetDecimal(0);
                if (dr.IsDBNull(1)) this.QuanChangLuoDiShiJiao = null; else this.QuanChangLuoDiShiJiao = dr.GetDecimal(1);
                find = true;
            }
            dr.Close();
            return find;
        }

        public bool GaoluSelect(OracleConnection cn,int gaolu, DateTime Rq)
        {
           bool find = false;
            OracleCommand selectCmd = new OracleCommand();
            selectCmd.Connection = cn;
            selectCmd.CommandText = "SELECT 高炉 FROM 原料消耗 WHERE (日期=:RQ AND 高炉=:GAOLU)";
            selectCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = Rq;
            selectCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = gaolu;

            OracleDataReader dr = selectCmd.ExecuteReader();
            if (dr.Read())
            {
                find = true;
            }
            return find;
  
        }

        public void SaveData(int gaolu, DateTime Rq)
        {
            OracleConnection cn = new OracleConnection();
            cn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
            cn.Open();
            if (GaoluSelect(cn, gaolu, Rq))
            {
                GaoluUpdate(cn, gaolu, Rq);
            }
            else
            {
                GaoluInsert(cn, gaolu, Rq);
            }

            //if (RiSelect(cn, Rq))
            //{
            //    RiUpdate(cn, Rq);
            //}
            //else
            //{
            //    RiInsert(cn, Rq);
            //}
            cn.Close();
        }

        public void GaoluInsert(OracleConnection cn, int gaolu, DateTime Rq)
        {
            OracleCommand insertCmd = new OracleCommand();
            insertCmd.Connection = cn;
            insertCmd.CommandText = "INSERT INTO 原料消耗(日期,高炉,机烧矿, 球团矿, 国内球团矿,进口球团矿,PB块,纽曼块,其它块矿,高钛球团矿,高品位钛球,硅石,萤石,蛇纹石,其它熔剂,富氧量,工艺称,煤粉,焦丁,自产湿焦,落地湿焦,生成标志) values（:日期,:高炉,:机烧矿,:球团矿,:国内球团矿,:进口球团矿,:PB块,:纽曼块,:其它块矿,:高钛球团矿,:高品位钛球,:硅石,:萤石,:蛇纹石,:其它熔剂,:富氧量,:工艺称,:煤粉,:焦丁,:自产湿焦,:落地湿焦,1）";
            insertCmd.Parameters.Add(":日期", OracleType.DateTime, 7).Value = Rq;
            insertCmd.Parameters.Add(":高炉", OracleType.Int32, 22).Value = gaolu;
            insertCmd.Parameters.Add(":机烧矿", OracleType.Double, 22).Value = (object)this.ShaoJieKuang ?? DBNull.Value;
            insertCmd.Parameters.Add(":球团矿", OracleType.Double, 22).Value = (object)this.Qiutuankuang ?? DBNull.Value;
            insertCmd.Parameters.Add(":国内球团矿", OracleType.Double, 22).Value = (object)this.Guoneiqiutuankuang?? DBNull.Value;
            insertCmd.Parameters.Add(":进口球团矿", OracleType.Double, 22).Value = (object)this.Jinkouqiutuankuang ?? DBNull.Value;
            insertCmd.Parameters.Add(":PB块", OracleType.Double, 22).Value = (object)this.PBKuai ?? DBNull.Value;
            insertCmd.Parameters.Add(":纽曼块", OracleType.Double, 22).Value = (object)this.Niumankuai ?? DBNull.Value;
            insertCmd.Parameters.Add(":其它块矿", OracleType.Double, 22).Value = (object)this.Qitakuaikuang ?? DBNull.Value;

            insertCmd.Parameters.Add(":高钛球团矿", OracleType.Double, 22).Value = (object)this.Gaotaiqiutuankuang ?? DBNull.Value;
            insertCmd.Parameters.Add(":高品位钛球", OracleType.Double, 22).Value = (object)this.Gaopinweitaiqiu?? DBNull.Value;
            insertCmd.Parameters.Add(":硅石", OracleType.Double, 22).Value = (object)this.GuiShi ?? DBNull.Value;
            insertCmd.Parameters.Add(":萤石", OracleType.Double, 22).Value = (object)this.YingShi ?? DBNull.Value;
            insertCmd.Parameters.Add(":蛇纹石", OracleType.Double, 22).Value = (object)this.SheWenShi ?? DBNull.Value;
            insertCmd.Parameters.Add(":其它熔剂", OracleType.Double, 22).Value = (object)this.Qitarongji ?? DBNull.Value;
            insertCmd.Parameters.Add(":富氧量", OracleType.Double, 22).Value = (object)this.FuYangLiang ?? DBNull.Value;

            insertCmd.Parameters.Add(":工艺称", OracleType.Double, 22).Value = (object)this.GongYiCheng ?? DBNull.Value;
            insertCmd.Parameters.Add(":煤粉", OracleType.Double, 22).Value = (object)this.PenMei ?? DBNull.Value;
            insertCmd.Parameters.Add(":焦丁", OracleType.Double, 22).Value = (object)this.JiaoDing ?? DBNull.Value;
            
            
            insertCmd.Parameters.Add(":自产湿焦", OracleType.Double, 22).Value = (object)this.ZiChanShiJiao ?? DBNull.Value;
            insertCmd.Parameters.Add(":落地湿焦", OracleType.Double, 22).Value = (object)this.LuoDiShiJiao ?? DBNull.Value;
        
          
            

            insertCmd.ExecuteNonQuery();
        }

        public void GaoluUpdate(OracleConnection cn, int gaolu, DateTime Rq)
        {
            try
            {
                OracleCommand updateCmd = new OracleCommand();
                updateCmd.Connection = cn;
                // 机烧矿, 球团矿, 国内球团矿,进口球团矿,PB块,纽曼块,其它块矿,高钛球团矿,高品位钛球,硅石,萤石,蛇纹石,其它熔剂,富氧量,工艺称,煤粉,焦丁,自产湿焦,落地湿焦
                updateCmd.CommandText = "UPDATE 原料消耗 SET 机烧矿=:机烧矿,球团矿=:球团矿,国内球团矿=:国内球团矿,进口球团矿=:进口球团矿,PB块=:PB块,纽曼块=:纽曼块,其它块矿=:其它块矿,高钛球团矿=:高钛球团矿,高品位钛球=:高品位钛球,硅石=:硅石,萤石=:萤石,蛇纹石=:蛇纹石,其它熔剂=:其它熔剂,富氧量=:富氧量,工艺称=:工艺称,焦丁=:焦丁,煤粉=:煤粉,自产湿焦=:自产湿焦,落地湿焦=:落地湿焦 WHERE 日期=:日期 and 高炉=:高炉";

                updateCmd.Parameters.Add(":机烧矿", OracleType.Double, 22).Value = (object)this.ShaoJieKuang ?? DBNull.Value;
                updateCmd.Parameters.Add(":球团矿", OracleType.Double, 22).Value = (object)this.Qiutuankuang ?? DBNull.Value;
                updateCmd.Parameters.Add(":国内球团矿", OracleType.Double, 22).Value = (object)this.Guoneiqiutuankuang ?? DBNull.Value;
                updateCmd.Parameters.Add(":进口球团矿", OracleType.Double, 22).Value = (object)this.Jinkouqiutuankuang ?? DBNull.Value;
                updateCmd.Parameters.Add(":PB块", OracleType.Double, 22).Value = (object)this.PBKuai ?? DBNull.Value;
                updateCmd.Parameters.Add(":纽曼块", OracleType.Double, 22).Value = (object)this.Niumankuai ?? DBNull.Value;
                updateCmd.Parameters.Add(":其它块矿", OracleType.Double, 22).Value = (object)this.Qitakuaikuang ?? DBNull.Value;

                updateCmd.Parameters.Add(":高钛球团矿", OracleType.Double, 22).Value = (object)this.Gaotaiqiutuankuang ?? DBNull.Value;
                updateCmd.Parameters.Add(":高品位钛球", OracleType.Double, 22).Value = (object)this.Gaopinweitaiqiu ?? DBNull.Value;
                updateCmd.Parameters.Add(":硅石", OracleType.Double, 22).Value = (object)this.GuiShi ?? DBNull.Value;
                updateCmd.Parameters.Add(":萤石", OracleType.Double, 22).Value = (object)this.YingShi ?? DBNull.Value;
                updateCmd.Parameters.Add(":蛇纹石", OracleType.Double, 22).Value = (object)this.SheWenShi ?? DBNull.Value;
                updateCmd.Parameters.Add(":其它熔剂", OracleType.Double, 22).Value = (object)this.Qitarongji ?? DBNull.Value;
                updateCmd.Parameters.Add(":富氧量", OracleType.Double, 22).Value = (object)this.FuYangLiang ?? DBNull.Value;

                updateCmd.Parameters.Add(":工艺称", OracleType.Double, 22).Value = (object)this.GongYiCheng ?? DBNull.Value;
                updateCmd.Parameters.Add(":煤粉", OracleType.Double, 22).Value = (object)this.PenMei ?? DBNull.Value;
                updateCmd.Parameters.Add(":焦丁", OracleType.Double, 22).Value = (object)this.JiaoDing ?? DBNull.Value;


                updateCmd.Parameters.Add(":自产湿焦", OracleType.Double, 22).Value = (object)this.ZiChanShiJiao ?? DBNull.Value;
                updateCmd.Parameters.Add(":落地湿焦", OracleType.Double, 22).Value = (object)this.LuoDiShiJiao ?? DBNull.Value;
                updateCmd.Parameters.Add(":日期", OracleType.DateTime, 7).Value = Rq;
                updateCmd.Parameters.Add(":高炉", OracleType.Int32, 22).Value = gaolu;
                updateCmd.ExecuteNonQuery();

            }
            catch
            {

                return;
            }

           
        }
       
        public bool RiSelect(OracleConnection cn, DateTime Rq)
        {
            bool find = false;
            OracleCommand selectCmd = new OracleCommand();
            selectCmd.Connection = cn;
            selectCmd.CommandText = "SELECT 自产湿焦,落地湿焦 FROM 全厂日消耗 WHERE P01日期=:RQ";
            selectCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = Rq;
            OracleDataReader dr= selectCmd.ExecuteReader();
            if (dr.Read())
            {
                find = true;
            }
            dr.Close();
            return find;
        }

        public void RiInsert(OracleConnection cn,DateTime Rq)
        {
             OracleCommand insertCmd = new OracleCommand();
             insertCmd.Connection = cn;
             insertCmd.CommandText="INSERT INTO 全厂日消耗(P01日期,自产湿焦,落地湿焦) VALUES (:RQ,:ZICHAN,:LUODI)";
             insertCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = Rq;
             insertCmd.Parameters.Add(":ZICHAN", OracleType.Double, 22).Value = (object)this.QuanChangZiChanShiJiao ?? DBNull.Value;
             insertCmd.Parameters.Add(":LUODI", OracleType.Double, 22).Value = (object)this.QuanChangLuoDiShiJiao ?? DBNull.Value;
             insertCmd.ExecuteNonQuery();
        }

        public void RiUpdate(OracleConnection cn,DateTime Rq)
        {
            OracleCommand updateCmd = new OracleCommand();
             updateCmd.Connection = cn;
             updateCmd.CommandText = "UPDATE 全厂日消耗 SET 自产湿焦=:ZICHAN,落地湿焦=:LUODI WHERE P01日期=:RQ";
             updateCmd.Parameters.Add(":ZICHAN", OracleType.Double, 22).Value = (object)this.QuanChangZiChanShiJiao ?? DBNull.Value;
             updateCmd.Parameters.Add(":LUODI", OracleType.Double, 22).Value = (object)this.QuanChangLuoDiShiJiao ?? DBNull.Value; 
             updateCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = Rq;
             updateCmd.ExecuteNonQuery();        
        }
    }
}
