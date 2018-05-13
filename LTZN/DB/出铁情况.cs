using System;
using System.Data;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.OracleClient;
using System.Reflection;

namespace ZHCDB
{
    //出铁
    public partial class Ddluci : DbEntity
    {
        private string _LUCI = "";
        private int _GAOLU;
        private System.DateTime _ZDSJ = DateTime.Now;
        private System.Nullable<System.DateTime> _DGSJ = null;
        private System.Nullable<System.DateTime> _KKSJ = null;
        private System.Nullable<System.DateTime> _DKSJ = null;
        private System.Nullable<double> _WDSJ;
        private string _QUCHU = "炼钢";
        private System.Nullable<double> _GFELIANG;
        private System.Nullable<double> _LFELIANG;
        private System.Nullable<double> _FELIANG;
        private System.Nullable<double> _LIAOPI;
        private System.Nullable<double> _GGUANSHU;
        private System.Nullable<double> _FEWENDU;
        private System.Nullable<double> _FEC;
        private System.Nullable<double> _FESI;
        private System.Nullable<double> _FEMN;
        private System.Nullable<double> _FEP;
        private System.Nullable<double> _FES;
        private System.Nullable<double> _FETI;
        private System.Nullable<double> _GFESI;
        private System.Nullable<double> _GFES;
        private System.Nullable<double> _SHENDU;
        private System.Nullable<double> _JIAODU;
        private System.Nullable<double> _DANILIANG;
        private string _ZHAYANG = "";
        private System.Nullable<double> _SHUIWEN1;
        private System.Nullable<double> _SHUIWEN2;
        private System.Nullable<double> _ZHASIO2;
        private System.Nullable<double> _ZHACAO;
        private System.Nullable<double> _ZHAMGO;
        private System.Nullable<double> _ZHAAL2O3;
        private System.Nullable<double> _ZHAS;
        private System.Nullable<double> _ZHATIO2;
        private System.Nullable<double> _ZHAR2;
        private string _BANCI = "";
        private int _BANLUCI;
        private System.Nullable<System.DateTime> _TZSJ = null;
        private string _BZ = "";
        private string _PNQK1 = "";
        private string _PNQK2 = "";
        private System.Nullable<double> _ZHAR3;
        private System.Nullable<double> _ZHAR4;
        private System.Nullable<double> _ZHAR5;
        private string _CTC = "";
        private System.Nullable<System.DateTime> _LZQK = null;
        //ConStr 

        [Query]
        [DisplayName("炉次")]
        public string LUCI
        {
            get
            {
                return this._LUCI;
            }
            set
            {
                if (!_LUCI.Equals(value))
                {
                    _LUCI = value;
                    RaisePropertyChanged("LUCI", true);
                }

            }
        }

        [Query]
        [DisplayName("高炉")]
        public int GAOLU
        {
            get
            {
                return this._GAOLU;
            }
            set
            {
                if (!_GAOLU.Equals(value))
                {
                    _GAOLU = value;
                    RaisePropertyChanged("GAOLU", true);
                }

            }
        }

        [Query]
        [DisplayName("正点时间")]
        public System.DateTime ZDSJ
        {
            get
            {
                return this._ZDSJ;
            }
            set
            {
                if (!_ZDSJ.Equals(value))
                {
                    _ZDSJ = value;
                    RaisePropertyChanged("ZDSJ", true);
                }

            }
        }

        private int _RIXUHAO;
        [DisplayName("日序号")]
        public int RIXUHAO
        {
            get { return _RIXUHAO; }
            set
            {
                if (_RIXUHAO != value)
                {
                    _RIXUHAO = value;
                    RaisePropertyChanged("RIXUHAO");
                }
            }
        }

        public string JYLUCI
        {
            get { return GAOLU.ToString() + RIXUHAO.ToString("00"); }
        }

        [Query]
        [DisplayName("开口时间")]
        public System.Nullable<System.DateTime> KKSJ
        {
            get
            {
                return this._KKSJ;
            }
            set
            {
                DateTime? val = XiuZheng(value);
                if (!_KKSJ.Equals(val))
                {
                    _KKSJ = val;
                    RaisePropertyChanged("KKSJ", true);
                    RaisePropertyChanged("CTJG", true);
                }

            }
        }

        [Query]
        [DisplayName("堵口时间")]
        public System.Nullable<System.DateTime> DKSJ
        {
            get
            {
                return this._DKSJ;
            }
            set
            {
                DateTime? val = XiuZheng(value);
                if (!_DKSJ.Equals(val))
                {
                    _DKSJ = val;
                    RaisePropertyChanged("DKSJ", true);
                    RaisePropertyChanged("CTJG", true);
                    CalWdsj();
                }

            }
        }

        [DisplayName("出铁间隔")]
        public int? CTJG
        {
            get
            {
                int? result = null;
                if (DKSJ != null && KKSJ != null)
                {
                    TimeSpan ctjg = DKSJ.Value - KKSJ.Value;
                    result = Convert.ToInt32(ctjg.TotalMinutes);
                }
                return result;
            }
        }

        [Query]
        [DisplayName("对罐时间")]
        public System.Nullable<System.DateTime> DGSJ
        {
            get
            {
                return this._DGSJ;
            }
            set
            {
                DateTime? val = XiuZheng(value);
                if (!_DGSJ.Equals(val))
                {
                    _DGSJ = val;
                    RaisePropertyChanged("DGSJ", true);
                    CalWdsj();
                }

            }
        }

        private DateTime? _GZSJ = null;

        [Query]
        [DisplayName("罐走时间")]
        public System.Nullable<System.DateTime> GZSJ
        {
            get
            {
                return this._GZSJ;
            }
            set
            {
                DateTime? val = XiuZheng(value);
                if (!_GZSJ.Equals(val))
                {
                    _GZSJ = val;
                    RaisePropertyChanged("GZSJ", true);
                }
            }
        }

        private int? _GZJG = null;

        [DisplayName("罐走间隔")]
        public int? GZJG
        {
            get
            {
                return this._GZJG;
            }
            set
            {
                if (!_GZJG.Equals(value))
                {
                    _GZJG = value;
                    RaisePropertyChanged("GZJG", true);
                }
            }
        }

        [Query]
        [DisplayName("通知时间")]
        public System.Nullable<System.DateTime> TZSJ
        {
            get
            {
                return this._TZSJ;
            }
            set
            {
                DateTime? val = XiuZheng(value);
                if (!_TZSJ.Equals(val))
                {
                    _TZSJ = val;
                    RaisePropertyChanged("TZSJ", true);
                }

            }
        }

        [Query]
        [DisplayName("晚点时间")]
        public System.Nullable<double> WDSJ
        {
            get
            {
                return this._WDSJ;
            }
            set
            {
                if (!_WDSJ.Equals(value))
                {
                    _WDSJ = value;
                    RaisePropertyChanged("WDSJ", true);
                }

            }
        }

        [Query]
        [DisplayName("去处")]
        public string QUCHU
        {
            get
            {
                return this._QUCHU;
            }
            set
            {
                if (!_QUCHU.Equals(value))
                {
                    _QUCHU = value;
                    RaisePropertyChanged("QUCHU", true);
                }

            }
        }

        [Query]
        [DisplayName("估计铁量")]
        public System.Nullable<double> GFELIANG
        {
            get
            {
                return this._GFELIANG;
            }
            set
            {
                if (!_GFELIANG.Equals(value))
                {
                    _GFELIANG = value;
                    RaisePropertyChanged("GFELIANG", true);
                }

            }
        }

        [Query]
        [DisplayName("理论铁量")]
        public System.Nullable<double> LFELIANG
        {
            get
            {
                return this._LFELIANG;
            }
            set
            {
                if (!_LFELIANG.Equals(value))
                {
                    _LFELIANG = value;
                    RaisePropertyChanged("LFELIANG", true);
                    RaisePropertyChanged("PIANCHA");
                }

            }
        }

        private double? _PIANCHA;
        [DisplayName("铁量差")]
        public double? PIANCHA
        {
            get
            {
                _PIANCHA = (this.FELIANG - this.LFELIANG);
                return _PIANCHA;
            }
            set
            {
                if (_PIANCHA != value)
                {
                    _PIANCHA = value;
                    RaisePropertyChanged("PIANCHA");
                }
            }
        }

        [Query]
        [DisplayName("实际铁量")]
        public System.Nullable<double> FELIANG
        {
            get
            {
                return this._FELIANG;
            }
            set
            {
                if (!_FELIANG.Equals(value))
                {
                    _FELIANG = value;
                    RaisePropertyChanged("FELIANG", true);
                    RaisePropertyChanged("PIANCHA");
                }

            }
        }

        [Query]
        [DisplayName("料批数")]
        public System.Nullable<double> LIAOPI
        {
            get
            {
                return this._LIAOPI;
            }
            set
            {
                if (!_LIAOPI.Equals(value))
                {
                    _LIAOPI = value;
                    RaisePropertyChanged("LIAOPI", true);
                }

            }
        }

        [Query]
        [DisplayName("估计罐数")]
        public System.Nullable<double> GGUANSHU
        {
            get
            {
                return this._GGUANSHU;
            }
            set
            {
                if (!_GGUANSHU.Equals(value))
                {
                    _GGUANSHU = value;
                    RaisePropertyChanged("GGUANSHU", true);
                }

            }
        }

        [Query]
        [DisplayName("铁水温度")]
        public System.Nullable<double> FEWENDU
        {
            get
            {
                return this._FEWENDU;
            }
            set
            {
                if (!_FEWENDU.Equals(value))
                {
                    _FEWENDU = value;
                    RaisePropertyChanged("FEWENDU", true);
                }

            }
        }

        [Query]
        [DisplayName("铁水温度")]
        public System.Nullable<double> FEC
        {
            get
            {
                return this._FEC;
            }
            set
            {
                if (!_FEC.Equals(value))
                {
                    _FEC = value;
                    RaisePropertyChanged("FEC", true);
                }

            }
        }

        [Query]
        [DisplayName("铁水Si")]
        public System.Nullable<double> FESI
        {
            get
            {
                return this._FESI;
            }
            set
            {
                if (!_FESI.Equals(value))
                {
                    _FESI = value;
                    RaisePropertyChanged("FESI", true);
                }

            }
        }

        [Query]
        [DisplayName("铁水Mn")]
        public System.Nullable<double> FEMN
        {
            get
            {
                return this._FEMN;
            }
            set
            {
                if (!_FEMN.Equals(value))
                {
                    _FEMN = value;
                    RaisePropertyChanged("FEMN", true);
                }

            }
        }

        [Query]
        [DisplayName("铁水P")]
        public System.Nullable<double> FEP
        {
            get
            {
                return this._FEP;
            }
            set
            {
                if (!_FEP.Equals(value))
                {
                    _FEP = value;
                    RaisePropertyChanged("FEP", true);
                }

            }
        }

        [Query]
        [DisplayName("铁水S")]
        public System.Nullable<double> FES
        {
            get
            {
                return this._FES;
            }
            set
            {
                if (!_FES.Equals(value))
                {
                    _FES = value;
                    RaisePropertyChanged("FES", true);
                }

            }
        }

        [Query]
        [DisplayName("铁水Ti")]
        public System.Nullable<double> FETI
        {
            get
            {
                return this._FETI;
            }
            set
            {
                if (!_FETI.Equals(value))
                {
                    _FETI = value;
                    RaisePropertyChanged("FETI", true);
                }

            }
        }

        [Query]
        [DisplayName("估计Si")]
        public System.Nullable<double> GFESI
        {
            get
            {
                return this._GFESI;
            }
            set
            {
                if (!_GFESI.Equals(value))
                {
                    _GFESI = value;
                    RaisePropertyChanged("GFESI", true);
                }

            }
        }

        [Query]
        [DisplayName("估计S")]
        public System.Nullable<double> GFES
        {
            get
            {
                return this._GFES;
            }
            set
            {
                if (!_GFES.Equals(value))
                {
                    _GFES = value;
                    RaisePropertyChanged("GFES", true);
                }

            }
        }

        [Query]
        [DisplayName("深度")]
        public System.Nullable<double> SHENDU
        {
            get
            {
                return this._SHENDU;
            }
            set
            {
                if (!_SHENDU.Equals(value))
                {
                    _SHENDU = value;
                    RaisePropertyChanged("SHENDU", true);
                }

            }
        }

        [Query]
        [DisplayName("角度")]
        public System.Nullable<double> JIAODU
        {
            get
            {
                return this._JIAODU;
            }
            set
            {
                if (!_JIAODU.Equals(value))
                {
                    _JIAODU = value;
                    RaisePropertyChanged("JIAODU", true);
                }

            }
        }

        [Query]
        [DisplayName("打泥量")]
        public System.Nullable<double> DANILIANG
        {
            get
            {
                return this._DANILIANG;
            }
            set
            {
                if (!_DANILIANG.Equals(value))
                {
                    _DANILIANG = value;
                    RaisePropertyChanged("DANILIANG", true);
                }

            }
        }

        [Query]
        [DisplayName("渣样")]
        public string ZHAYANG
        {
            get
            {
                return this._ZHAYANG;
            }
            set
            {
                if (!_ZHAYANG.Equals(value))
                {
                    _ZHAYANG = value;
                    RaisePropertyChanged("ZHAYANG", true);
                }

            }
        }

        [Query]
        [DisplayName("水温1")]
        public System.Nullable<double> SHUIWEN1
        {
            get
            {
                return this._SHUIWEN1;
            }
            set
            {
                if (!_SHUIWEN1.Equals(value))
                {
                    _SHUIWEN1 = value;
                    RaisePropertyChanged("SHUIWEN1", true);
                }

            }
        }

        [Query]
        [DisplayName("水温2")]
        public System.Nullable<double> SHUIWEN2
        {
            get
            {
                return this._SHUIWEN2;
            }
            set
            {
                if (!_SHUIWEN2.Equals(value))
                {
                    _SHUIWEN2 = value;
                    RaisePropertyChanged("SHUIWEN2", true);
                }

            }
        }

        [Query]
        [DisplayName("渣SiO2")]
        public System.Nullable<double> ZHASIO2
        {
            get
            {
                return this._ZHASIO2;
            }
            set
            {
                if (!_ZHASIO2.Equals(value))
                {
                    _ZHASIO2 = value;
                    RaisePropertyChanged("ZHASIO2", true);
                }

            }
        }

        [Query]
        [DisplayName("渣CaO")]
        public System.Nullable<double> ZHACAO
        {
            get
            {
                return this._ZHACAO;
            }
            set
            {
                if (!_ZHACAO.Equals(value))
                {
                    _ZHACAO = value;
                    RaisePropertyChanged("ZHACAO", true);
                }

            }
        }

        [Query]
        [DisplayName("渣MgO")]
        public System.Nullable<double> ZHAMGO
        {
            get
            {
                return this._ZHAMGO;
            }
            set
            {
                if (!_ZHAMGO.Equals(value))
                {
                    _ZHAMGO = value;
                    RaisePropertyChanged("ZHAMGO", true);
                }

            }
        }

        [Query]
        [DisplayName("渣Al2O3")]
        public System.Nullable<double> ZHAAL2O3
        {
            get
            {
                return this._ZHAAL2O3;
            }
            set
            {
                if (!_ZHAAL2O3.Equals(value))
                {
                    _ZHAAL2O3 = value;
                    RaisePropertyChanged("ZHAAL2O3", true);
                }

            }
        }

        [Query]
        [DisplayName("渣S")]
        public System.Nullable<double> ZHAS
        {
            get
            {
                return this._ZHAS;
            }
            set
            {
                if (!_ZHAS.Equals(value))
                {
                    _ZHAS = value;
                    RaisePropertyChanged("ZHAS", true);
                }

            }
        }

        [Query]
        [DisplayName("渣TiO2")]
        public System.Nullable<double> ZHATIO2
        {
            get
            {
                return this._ZHATIO2;
            }
            set
            {
                if (!_ZHATIO2.Equals(value))
                {
                    _ZHATIO2 = value;
                    RaisePropertyChanged("ZHATIO2", true);
                }

            }
        }

        [Query]
        [DisplayName("渣R2")]
        public System.Nullable<double> ZHAR2
        {
            get
            {
                return this._ZHAR2;
            }
            set
            {
                if (!_ZHAR2.Equals(value))
                {
                    _ZHAR2 = value;
                    RaisePropertyChanged("ZHAR2", true);
                }

            }
        }

        [Query]
        [DisplayName("班次")]
        public string BANCI
        {
            get
            {
                return this._BANCI;
            }
            set
            {
                if (!_BANCI.Equals(value))
                {
                    _BANCI = value;
                    RaisePropertyChanged("BANCI", true);
                }

            }
        }

        [Query]
        [DisplayName("班次号")]
        public int BANLUCI
        {
            get
            {
                return this._BANLUCI;
            }
            set
            {
                if (!_BANLUCI.Equals(value))
                {
                    _BANLUCI = value;
                    RaisePropertyChanged("BANLUCI", true);
                }

            }
        }

        [DisplayName("备注")]
        public string BZ
        {
            get
            {
                return this._BZ;
            }
            set
            {
                if (!_BZ.Equals(value))
                {
                    _BZ = value;
                    RaisePropertyChanged("BZ", true);
                }

            }
        }

        [DisplayName("炮泥情况1")]
        public string PNQK1
        {
            get
            {
                return this._PNQK1;
            }
            set
            {
                if (!_PNQK1.Equals(value))
                {
                    _PNQK1 = value;
                    RaisePropertyChanged("PNQK1", true);
                }

            }
        }

        [DisplayName("炮泥情况2")]
        public string PNQK2
        {
            get
            {
                return this._PNQK2;
            }
            set
            {
                if (!_PNQK2.Equals(value))
                {
                    _PNQK2 = value;
                    RaisePropertyChanged("PNQK2", true);
                }

            }
        }

        [Query]
        [DisplayName("渣R3")]
        public System.Nullable<double> ZHAR3
        {
            get
            {
                return this._ZHAR3;
            }
            set
            {
                if (!_ZHAR3.Equals(value))
                {
                    _ZHAR3 = value;
                    RaisePropertyChanged("ZHAR3", true);
                }

            }
        }

        [Query]
        [DisplayName("渣R4")]
        public System.Nullable<double> ZHAR4
        {
            get
            {
                return this._ZHAR4;
            }
            set
            {
                if (!_ZHAR4.Equals(value))
                {
                    _ZHAR4 = value;
                    RaisePropertyChanged("ZHAR4", true);
                }

            }
        }

        [Query]
        [DisplayName("渣MgO/Al2O3")]
        public System.Nullable<double> ZHAR5
        {
            get
            {
                return this._ZHAR5;
            }
            set
            {
                if (!_ZHAR5.Equals(value))
                {
                    _ZHAR5 = value;
                    RaisePropertyChanged("ZHAR5", true);
                }

            }
        }

        [Query]
        [DisplayName("出铁场")]
        public string CTC
        {
            get
            {
                return this._CTC;
            }
            set
            {
                if (!_CTC.Equals(value))
                {
                    _CTC = value;
                    RaisePropertyChanged("CTC", true);
                }

            }
        }

        [Query]
        [DisplayName("来渣")]
        public System.Nullable<System.DateTime> LZQK
        {
            get
            {
                return this._LZQK;
            }
            set
            {
                DateTime? val = XiuZheng(value);
                if (!_LZQK.Equals(val))
                {
                    _LZQK = val;
                    RaisePropertyChanged("LZQK", true);
                }

            }
        }

        private string _BANZU = "";

        [Query]
        [DisplayName("班组")]
        public string BANZU
        {
            get { return _BANZU; }
            set
            {
                if (_BANZU != value)
                {
                    _BANZU = value;
                    RaisePropertyChanged("BANZU");
                }
            }
        }

        private string _GONGZHANG = "";

        [DisplayName("工长")]
        public string GONGZHANG
        {
            get { return _GONGZHANG; }
            set
            {
                if (_GONGZHANG != value)
                {
                    _GONGZHANG = value;
                    RaisePropertyChanged("GONGZHANG");
                }
            }
        }

        private double? _BBCL;
        public double? BBCL
        {
            get
            {
                return _BBCL;
            }
            set
            {
                if (!_BBCL.Equals(value))
                {
                    _BBCL = value;
                    RaisePropertyChanged("BBCL", true);
                }
            }
        }
        public bool CheckPk(OracleTransaction trans)
        {
            OracleConnection con = trans.Connection;
            OracleCommand checkCmd = new OracleCommand();
            checkCmd.Connection = con;
            checkCmd.Transaction = trans;
            checkCmd.CommandText = "SELECT ROWID FROM DDLUCI WHERE GAOLU=:GAOLU AND ZDSJ=:ZDSJ";
            checkCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = this.GAOLU;
            checkCmd.Parameters.Add(":ZDSJ", OracleType.DateTime).Value = this.ZDSJ;
            OracleDataReader dr = checkCmd.ExecuteReader();
            bool result = false;
            if (dr.Read())
            {
                OracleString rowid = dr.GetOracleString(0);
                if (!rowid.IsNull)
                {
                    this.RowId = rowid;
                    result = true;
                }
            }
            dr.Close();
            return result;
        }

        public override void DbInsert(OracleTransaction trans)
        {
            //InsertSql,InsertPara,AutoSql
            if (DataState == DataStateType.Add)
            {
                OracleConnection con = trans.Connection;
                if (CheckPk(trans))
                {
                    OracleCommand updateCmd = new OracleCommand();
                    updateCmd.Connection = con;
                    updateCmd.Transaction = trans;
                    updateCmd.CommandText = "UPDATE DDLUCI SET LUCI=:LUCI,DGSJ=:DGSJ,KKSJ=:KKSJ,DKSJ=:DKSJ,TZSJ=:TZSJ,WDSJ=:WDSJ,QUCHU=:QUCHU,GFELIANG=:GFELIANG,LFELIANG=:LFELIANG,PIANCHA=:PIANCHA,FELIANG=:FELIANG,LIAOPI=:LIAOPI,GGUANSHU=:GGUANSHU,FEWENDU=:FEWENDU,GFESI=:GFESI,GFES=:GFES,SHENDU=:SHENDU,JIAODU=:JIAODU,DANILIANG=:DANILIANG,ZHAYANG=:ZHAYANG,SHUIWEN1=:SHUIWEN1,SHUIWEN2=:SHUIWEN2,PNQK1=:PNQK1,PNQK2=:PNQK2,CTC=:CTC,LZQK=:LZQK,BZ=:BZ,BBCL=:BBCL  WHERE GAOLU=:GAOLU AND ZDSJ=:ZDSJ";
                    updateCmd.Parameters.Add(":LUCI", OracleType.VarChar, 20).Value = this.LUCI;
                    updateCmd.Parameters.Add(":DGSJ", OracleType.DateTime, 7).Value = (object)this.DGSJ ?? DBNull.Value;
                    updateCmd.Parameters.Add(":KKSJ", OracleType.DateTime, 7).Value = (object)this.KKSJ ?? DBNull.Value;
                    updateCmd.Parameters.Add(":DKSJ", OracleType.DateTime, 7).Value = (object)this.DKSJ ?? DBNull.Value;
                    updateCmd.Parameters.Add(":TZSJ", OracleType.DateTime, 7).Value = (object)this.TZSJ ?? DBNull.Value;
                    updateCmd.Parameters.Add(":WDSJ", OracleType.Double, 22).Value = (object)this.WDSJ ?? DBNull.Value;
                    updateCmd.Parameters.Add(":QUCHU", OracleType.VarChar, 20).Value = this.QUCHU;
                    updateCmd.Parameters.Add(":GFELIANG", OracleType.Double, 22).Value = (object)this.GFELIANG ?? DBNull.Value;
                    updateCmd.Parameters.Add(":LFELIANG", OracleType.Double, 22).Value = (object)this.LFELIANG ?? DBNull.Value;
                    updateCmd.Parameters.Add(":PIANCHA", OracleType.Double, 22).Value = (object)this.PIANCHA ?? DBNull.Value;
                    updateCmd.Parameters.Add(":FELIANG", OracleType.Double, 22).Value = (object)this.FELIANG ?? DBNull.Value;
                    updateCmd.Parameters.Add(":LIAOPI", OracleType.Double, 22).Value = (object)this.LIAOPI ?? DBNull.Value;
                    updateCmd.Parameters.Add(":GGUANSHU", OracleType.Double, 22).Value = (object)this.GGUANSHU ?? DBNull.Value;
                    updateCmd.Parameters.Add(":FEWENDU", OracleType.Double, 22).Value = (object)this.FEWENDU ?? DBNull.Value;
                    updateCmd.Parameters.Add(":GFESI", OracleType.Double, 22).Value = (object)this.GFESI ?? DBNull.Value;
                    updateCmd.Parameters.Add(":GFES", OracleType.Double, 22).Value = (object)this.GFES ?? DBNull.Value;
                    updateCmd.Parameters.Add(":SHENDU", OracleType.Double, 22).Value = (object)this.SHENDU ?? DBNull.Value;
                    updateCmd.Parameters.Add(":JIAODU", OracleType.Double, 22).Value = (object)this.JIAODU ?? DBNull.Value;
                    updateCmd.Parameters.Add(":DANILIANG", OracleType.Double, 22).Value = (object)this.DANILIANG ?? DBNull.Value;
                    updateCmd.Parameters.Add(":ZHAYANG", OracleType.VarChar, 20).Value = this.ZHAYANG;
                    updateCmd.Parameters.Add(":SHUIWEN1", OracleType.Double, 22).Value = (object)this.SHUIWEN1 ?? DBNull.Value;
                    updateCmd.Parameters.Add(":SHUIWEN2", OracleType.Double, 22).Value = (object)this.SHUIWEN2 ?? DBNull.Value;
                    updateCmd.Parameters.Add(":PNQK1", OracleType.VarChar, 50).Value = this.PNQK1;
                    updateCmd.Parameters.Add(":PNQK2", OracleType.VarChar, 50).Value = this.PNQK2;
                    updateCmd.Parameters.Add(":CTC", OracleType.VarChar, 100).Value = this.CTC;
                    updateCmd.Parameters.Add(":LZQK", OracleType.DateTime, 7).Value = (object)this.LZQK ?? DBNull.Value;
                    updateCmd.Parameters.Add(":BZ", OracleType.VarChar, 50).Value = this.BZ;
                    updateCmd.Parameters.Add(":BBCL", OracleType.Double, 22).Value = (object)this.BBCL ?? DBNull.Value;
                    updateCmd.Parameters.Add(":GAOLU", OracleType.Int32, 22).Value = this.GAOLU;
                    updateCmd.Parameters.Add(":ZDSJ", OracleType.DateTime, 7).Value = this.ZDSJ;
                    updateCmd.ExecuteOracleNonQuery(out RowId);
                    return;
                }

                OracleCommand insertCmd = new OracleCommand();
                insertCmd.Connection = con;
                insertCmd.Transaction = trans;
                insertCmd.CommandText = "INSERT INTO DDLUCI(GAOLU,ZDSJ,LUCI,DGSJ,KKSJ,DKSJ,TZSJ,WDSJ,QUCHU,GFELIANG,LFELIANG,PIANCHA,FELIANG,LIAOPI,GGUANSHU,FEWENDU,FEC,FESI,FEMN,FEP,FES,FETI,GFESI,GFES,ZHASIO2,ZHACAO,ZHAMGO,ZHAAL2O3,ZHAS,ZHATIO2,ZHAR2,ZHAR3,ZHAR4,ZHAR5,SHENDU,JIAODU,DANILIANG,ZHAYANG,SHUIWEN1,SHUIWEN2,PNQK1,PNQK2,CTC,LZQK,BZ,BANCI,BANLUCI,BANZU,GONGZHANG,RIXUHAO,GZSJ,BBCL) VALUES(:GAOLU,:ZDSJ,:LUCI,:DGSJ,:KKSJ,:DKSJ,:TZSJ,:WDSJ,:QUCHU,:GFELIANG,:LFELIANG,:PIANCHA,:FELIANG,:LIAOPI,:GGUANSHU,:FEWENDU,:FEC,:FESI,:FEMN,:FEP,:FES,:FETI,:GFESI,:GFES,:ZHASIO2,:ZHACAO,:ZHAMGO,:ZHAAL2O3,:ZHAS,:ZHATIO2,:ZHAR2,:ZHAR3,:ZHAR4,:ZHAR5,:SHENDU,:JIAODU,:DANILIANG,:ZHAYANG,:SHUIWEN1,:SHUIWEN2,:PNQK1,:PNQK2,:CTC,:LZQK,:BZ,:BANCI,:BANLUCI,:BANZU,:GONGZHANG,:RIXUHAO,:GZSJ,:BBCL)";
                insertCmd.Parameters.Add(":GAOLU", OracleType.Int32, 22).Value = this.GAOLU;
                insertCmd.Parameters.Add(":ZDSJ", OracleType.DateTime, 7).Value = this.ZDSJ;
                insertCmd.Parameters.Add(":LUCI", OracleType.VarChar, 20).Value = this.LUCI;
                insertCmd.Parameters.Add(":DGSJ", OracleType.DateTime, 7).Value = (object)this.DGSJ ?? DBNull.Value;
                insertCmd.Parameters.Add(":KKSJ", OracleType.DateTime, 7).Value = (object)this.KKSJ ?? DBNull.Value;
                insertCmd.Parameters.Add(":DKSJ", OracleType.DateTime, 7).Value = (object)this.DKSJ ?? DBNull.Value;
                insertCmd.Parameters.Add(":TZSJ", OracleType.DateTime, 7).Value = (object)this.TZSJ ?? DBNull.Value;
                insertCmd.Parameters.Add(":WDSJ", OracleType.Double, 22).Value = (object)this.WDSJ ?? DBNull.Value;
                insertCmd.Parameters.Add(":QUCHU", OracleType.VarChar, 20).Value = this.QUCHU;
                insertCmd.Parameters.Add(":GFELIANG", OracleType.Double, 22).Value = (object)this.GFELIANG ?? DBNull.Value;
                insertCmd.Parameters.Add(":LFELIANG", OracleType.Double, 22).Value = (object)this.LFELIANG ?? DBNull.Value;
                insertCmd.Parameters.Add(":PIANCHA", OracleType.Double, 22).Value = (object)this.PIANCHA ?? DBNull.Value;
                insertCmd.Parameters.Add(":FELIANG", OracleType.Double, 22).Value = (object)this.FELIANG ?? DBNull.Value;
                insertCmd.Parameters.Add(":LIAOPI", OracleType.Double, 22).Value = (object)this.LIAOPI ?? DBNull.Value;
                insertCmd.Parameters.Add(":GGUANSHU", OracleType.Double, 22).Value = (object)this.GGUANSHU ?? DBNull.Value;
                insertCmd.Parameters.Add(":FEWENDU", OracleType.Double, 22).Value = (object)this.FEWENDU ?? DBNull.Value;
                insertCmd.Parameters.Add(":FEC", OracleType.Double, 22).Value = (object)this.FEC ?? DBNull.Value;
                insertCmd.Parameters.Add(":FESI", OracleType.Double, 22).Value = (object)this.FESI ?? DBNull.Value;
                insertCmd.Parameters.Add(":FEMN", OracleType.Double, 22).Value = (object)this.FEMN ?? DBNull.Value;
                insertCmd.Parameters.Add(":FEP", OracleType.Double, 22).Value = (object)this.FEP ?? DBNull.Value;
                insertCmd.Parameters.Add(":FES", OracleType.Double, 22).Value = (object)this.FES ?? DBNull.Value;
                insertCmd.Parameters.Add(":FETI", OracleType.Double, 22).Value = (object)this.FETI ?? DBNull.Value;
                insertCmd.Parameters.Add(":GFESI", OracleType.Double, 22).Value = (object)this.GFESI ?? DBNull.Value;
                insertCmd.Parameters.Add(":GFES", OracleType.Double, 22).Value = (object)this.GFES ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHASIO2", OracleType.Double, 22).Value = (object)this.ZHASIO2 ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHACAO", OracleType.Double, 22).Value = (object)this.ZHACAO ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHAMGO", OracleType.Double, 22).Value = (object)this.ZHAMGO ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHAAL2O3", OracleType.Double, 22).Value = (object)this.ZHAAL2O3 ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHAS", OracleType.Double, 22).Value = (object)this.ZHAS ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHATIO2", OracleType.Double, 22).Value = (object)this.ZHATIO2 ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHAR2", OracleType.Double, 22).Value = (object)this.ZHAR2 ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHAR3", OracleType.Double, 22).Value = (object)this.ZHAR3 ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHAR4", OracleType.Double, 22).Value = (object)this.ZHAR4 ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHAR5", OracleType.Double, 22).Value = (object)this.ZHAR5 ?? DBNull.Value;
                insertCmd.Parameters.Add(":SHENDU", OracleType.Double, 22).Value = (object)this.SHENDU ?? DBNull.Value;
                insertCmd.Parameters.Add(":JIAODU", OracleType.Double, 22).Value = (object)this.JIAODU ?? DBNull.Value;
                insertCmd.Parameters.Add(":DANILIANG", OracleType.Double, 22).Value = (object)this.DANILIANG ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHAYANG", OracleType.VarChar, 20).Value = this.ZHAYANG;
                insertCmd.Parameters.Add(":SHUIWEN1", OracleType.Double, 22).Value = (object)this.SHUIWEN1 ?? DBNull.Value;
                insertCmd.Parameters.Add(":SHUIWEN2", OracleType.Double, 22).Value = (object)this.SHUIWEN2 ?? DBNull.Value;
                insertCmd.Parameters.Add(":PNQK1", OracleType.VarChar, 50).Value = this.PNQK1;
                insertCmd.Parameters.Add(":PNQK2", OracleType.VarChar, 50).Value = this.PNQK2;
                insertCmd.Parameters.Add(":CTC", OracleType.VarChar, 100).Value = this.CTC;
                insertCmd.Parameters.Add(":LZQK", OracleType.DateTime, 7).Value = (object)this.LZQK ?? DBNull.Value;
                insertCmd.Parameters.Add(":BZ", OracleType.VarChar, 50).Value = this.BZ;
                insertCmd.Parameters.Add(":BANCI", OracleType.VarChar, 20).Value = this.BANCI;
                insertCmd.Parameters.Add(":BANLUCI", OracleType.Int32, 22).Value = this.BANLUCI;
                insertCmd.Parameters.Add(":BANZU", OracleType.VarChar, 50).Value = this.BANZU;
                insertCmd.Parameters.Add(":GONGZHANG", OracleType.VarChar, 50).Value = this.GONGZHANG;
                insertCmd.Parameters.Add(":RIXUHAO", OracleType.Int32).Value = this.RIXUHAO;
                insertCmd.Parameters.Add(":GZSJ", OracleType.DateTime).Value = (object)this.GZSJ ?? DBNull.Value;
                insertCmd.Parameters.Add(":BBCL", OracleType.Double, 22).Value = (object)this.BBCL ?? DBNull.Value;
                insertCmd.ExecuteOracleNonQuery(out RowId);
            }
        }
        public override void DbDelete(OracleTransaction trans)
        {
            //DelSql,DelPara
            if (DataState == DataStateType.Delete)
            {
                OracleConnection con = trans.Connection;
                OracleCommand deleteCmd = new OracleCommand();
                deleteCmd.Connection = con;
                deleteCmd.Transaction = trans;
                deleteCmd.CommandText = "DELETE FROM DDLUCI WHERE ROWID='" + RowId.Value + "'";

                deleteCmd.ExecuteNonQuery();
            }

        }
        public override void DbUpdate(OracleTransaction trans)
        {
            //UpdateSql,UpdatePara
            if (DataState == DataStateType.Update)
            {
                OracleConnection con = trans.Connection;
                OracleCommand updateCmd = new OracleCommand();
                updateCmd.Connection = con;
                updateCmd.Transaction = trans;
                updateCmd.CommandText = "UPDATE DDLUCI SET LUCI=:LUCI,DGSJ=:DGSJ,KKSJ=:KKSJ,DKSJ=:DKSJ,TZSJ=:TZSJ,WDSJ=:WDSJ,QUCHU=:QUCHU,GFELIANG=:GFELIANG,LFELIANG=:LFELIANG,PIANCHA=:PIANCHA,FELIANG=:FELIANG,LIAOPI=:LIAOPI,GGUANSHU=:GGUANSHU,FEWENDU=:FEWENDU,GFESI=:GFESI,GFES=:GFES,SHENDU=:SHENDU,JIAODU=:JIAODU,DANILIANG=:DANILIANG,ZHAYANG=:ZHAYANG,SHUIWEN1=:SHUIWEN1,SHUIWEN2=:SHUIWEN2,PNQK1=:PNQK1,PNQK2=:PNQK2,CTC=:CTC,LZQK=:LZQK,BZ=:BZ,GZSJ=:GZSJ,BBCL=:BBCL WHERE ROWID='" + RowId.Value + "'";
                updateCmd.Parameters.Add(":LUCI", OracleType.VarChar, 20).Value = this.LUCI;
                updateCmd.Parameters.Add(":DGSJ", OracleType.DateTime, 7).Value = (object)this.DGSJ ?? DBNull.Value;
                updateCmd.Parameters.Add(":KKSJ", OracleType.DateTime, 7).Value = (object)this.KKSJ ?? DBNull.Value;
                updateCmd.Parameters.Add(":DKSJ", OracleType.DateTime, 7).Value = (object)this.DKSJ ?? DBNull.Value;
                updateCmd.Parameters.Add(":TZSJ", OracleType.DateTime, 7).Value = (object)this.TZSJ ?? DBNull.Value;
                updateCmd.Parameters.Add(":WDSJ", OracleType.Double, 22).Value = (object)this.WDSJ ?? DBNull.Value;
                updateCmd.Parameters.Add(":QUCHU", OracleType.VarChar, 20).Value = this.QUCHU;
                updateCmd.Parameters.Add(":GFELIANG", OracleType.Double, 22).Value = (object)this.GFELIANG ?? DBNull.Value;
                updateCmd.Parameters.Add(":LFELIANG", OracleType.Double, 22).Value = (object)this.LFELIANG ?? DBNull.Value;
                updateCmd.Parameters.Add(":PIANCHA", OracleType.Double, 22).Value = (object)this.PIANCHA ?? DBNull.Value;
                updateCmd.Parameters.Add(":FELIANG", OracleType.Double, 22).Value = (object)this.FELIANG ?? DBNull.Value;
                updateCmd.Parameters.Add(":LIAOPI", OracleType.Double, 22).Value = (object)this.LIAOPI ?? DBNull.Value;
                updateCmd.Parameters.Add(":GGUANSHU", OracleType.Double, 22).Value = (object)this.GGUANSHU ?? DBNull.Value;
                updateCmd.Parameters.Add(":FEWENDU", OracleType.Double, 22).Value = (object)this.FEWENDU ?? DBNull.Value;
                updateCmd.Parameters.Add(":GFESI", OracleType.Double, 22).Value = (object)this.GFESI ?? DBNull.Value;
                updateCmd.Parameters.Add(":GFES", OracleType.Double, 22).Value = (object)this.GFES ?? DBNull.Value;
                updateCmd.Parameters.Add(":SHENDU", OracleType.Double, 22).Value = (object)this.SHENDU ?? DBNull.Value;
                updateCmd.Parameters.Add(":JIAODU", OracleType.Double, 22).Value = (object)this.JIAODU ?? DBNull.Value;
                updateCmd.Parameters.Add(":DANILIANG", OracleType.Double, 22).Value = (object)this.DANILIANG ?? DBNull.Value;
                updateCmd.Parameters.Add(":ZHAYANG", OracleType.VarChar, 20).Value = this.ZHAYANG;
                updateCmd.Parameters.Add(":SHUIWEN1", OracleType.Double, 22).Value = (object)this.SHUIWEN1 ?? DBNull.Value;
                updateCmd.Parameters.Add(":SHUIWEN2", OracleType.Double, 22).Value = (object)this.SHUIWEN2 ?? DBNull.Value;
                updateCmd.Parameters.Add(":PNQK1", OracleType.VarChar, 50).Value = this.PNQK1;
                updateCmd.Parameters.Add(":PNQK2", OracleType.VarChar, 50).Value = this.PNQK2;
                updateCmd.Parameters.Add(":CTC", OracleType.VarChar, 100).Value = this.CTC;
                updateCmd.Parameters.Add(":LZQK", OracleType.DateTime, 7).Value = (object)this.LZQK ?? DBNull.Value;
                updateCmd.Parameters.Add(":BZ", OracleType.VarChar, 50).Value = this.BZ;
                updateCmd.Parameters.Add(":GZSJ", OracleType.DateTime).Value = (object)this.GZSJ ?? DBNull.Value;
                updateCmd.Parameters.Add(":BBCL", OracleType.Double, 22).Value = (object)this.BBCL ?? DBNull.Value;
                updateCmd.ExecuteNonQuery();
            }
        }

        private DateTime? XiuZheng(DateTime? time)
        {
            if (time.HasValue)
            {
                DateTime result = new DateTime(ZDSJ.Year, ZDSJ.Month, ZDSJ.Day, time.Value.Hour, time.Value.Minute, 0);
                if (result > ZDSJ)
                {
                    TimeSpan span = result - ZDSJ;
                    if (span.TotalHours > 22)
                        result = result.AddDays(-1);

                }
                else
                {
                    TimeSpan span = ZDSJ - result;
                    if (span.TotalHours > 22)
                        result = result.AddDays(1);
                }
                return result;
            }
            else
                return null;
        }

        private void CalWdsj()
        {
            if (DKSJ != null)
            {
                TimeSpan jiange = DKSJ.Value - ZDSJ;
                if (DGSJ != null)
                {
                    if (ZDSJ < DGSJ.Value.AddMinutes(50))
                    {
                        jiange = DKSJ.Value - DGSJ.Value.AddMinutes(50);
                    }
                }
                if (jiange.TotalMinutes > 0)
                    WDSJ = jiange.TotalMinutes;
                else
                    WDSJ = 0;
            }
            else
                WDSJ = null;
        }

        public override bool Equals(object obj)
        {
            Ddluci other = obj as Ddluci;
            if (other != null)
            {
                if (this.GAOLU == other.GAOLU && this.ZDSJ == other.ZDSJ)
                    return true;
            }
            return false;
        }

        public override int GetHashCode()
        {
            // int hash = this.ZDSJ.Year * 10000 + this.ZDSJ.Month * 10000 + this.ZDSJ.Day * 1000 + this.ZDSJ.Hour * 100 + this.ZDSJ.Minute + GAOLU;
            return base.GetHashCode();
        }
    }

 
    public partial class DdluciTable : DbEntityTable<Ddluci>
    {

        public int lastLoadGaolu = 0;  //最后查询的高炉
        public DateTime? lastLoadRq = null; // 最后查询的日期

        public void LoadByGLRQ(int gaolu, DateTime rq)
        {
            lastLoadGaolu = gaolu;
            lastLoadRq = rq;
            lastEditGaolu = 0;
            lastEditZdsj = null;
            beginAdjustData = true;
            this.Clear();
            hideList.Clear();
            OracleConnection con = new OracleConnection(LTZN.Properties.Settings.Default.ltznConnectionString);
            con.Open();
            OracleCommand selectCmd = new OracleCommand();
            selectCmd.Connection = con;
            selectCmd.CommandText = "SELECT GAOLU,ZDSJ,LUCI,DGSJ,KKSJ,DKSJ,TZSJ,WDSJ,QUCHU,GFELIANG,LFELIANG,PIANCHA,FELIANG,LIAOPI,GGUANSHU,FEWENDU,FEC,FESI,FEMN,FEP,FES,FETI,GFESI,GFES,ZHASIO2,ZHACAO,ZHAMGO,ZHAAL2O3,ZHAS,ZHATIO2,ZHAR2,ZHAR3,ZHAR4,ZHAR5,SHENDU,JIAODU,DANILIANG,ZHAYANG,SHUIWEN1,SHUIWEN2,PNQK1,PNQK2,CTC,LZQK,BZ,BANCI,BANLUCI,BANZU,GONGZHANG,RIXUHAO,GZSJ,round((KKSJ-lag(DKSJ) over (order by ZDSJ))*1440,0) as GZJG,BBCL,ROWID FROM DDLUCI WHERE GAOLU=:GAOLU AND TRUNC(ZDSJ)=:RQ";
            selectCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = gaolu;
            selectCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = rq;

            OracleDataReader dr = selectCmd.ExecuteReader();
            while (dr.Read())
            {
                Ddluci item = new Ddluci();
                item.BeginInit();
                if (dr.IsDBNull(0)) item.GAOLU = 0; else item.GAOLU = dr.GetInt32(0);
                if (dr.IsDBNull(1)) item.ZDSJ = DateTime.Now; else item.ZDSJ = dr.GetDateTime(1);
                if (dr.IsDBNull(2)) item.LUCI = ""; else item.LUCI = dr.GetString(2);
                if (dr.IsDBNull(3)) item.DGSJ = null; else item.DGSJ = dr.GetDateTime(3);
                if (dr.IsDBNull(4)) item.KKSJ = null; else item.KKSJ = dr.GetDateTime(4);
                if (dr.IsDBNull(5)) item.DKSJ = null; else item.DKSJ = dr.GetDateTime(5);
                if (dr.IsDBNull(6)) item.TZSJ = null; else item.TZSJ = dr.GetDateTime(6);
                if (dr.IsDBNull(7)) item.WDSJ = null; else item.WDSJ = dr.GetDouble(7);
                if (dr.IsDBNull(8)) item.QUCHU = ""; else item.QUCHU = dr.GetString(8);
                if (dr.IsDBNull(9)) item.GFELIANG = null; else item.GFELIANG = dr.GetDouble(9);
                if (dr.IsDBNull(10)) item.LFELIANG = null; else item.LFELIANG = dr.GetDouble(10);
                if (dr.IsDBNull(11)) item.PIANCHA = null; else item.PIANCHA = dr.GetDouble(11);
                if (dr.IsDBNull(12)) item.FELIANG = null; else item.FELIANG = dr.GetDouble(12);
                if (dr.IsDBNull(13)) item.LIAOPI = null; else item.LIAOPI = dr.GetDouble(13);
                if (dr.IsDBNull(14)) item.GGUANSHU = null; else item.GGUANSHU = dr.GetDouble(14);
                if (dr.IsDBNull(15)) item.FEWENDU = null; else item.FEWENDU = dr.GetDouble(15);
                if (dr.IsDBNull(16)) item.FEC = null; else item.FEC = dr.GetDouble(16);
                if (dr.IsDBNull(17)) item.FESI = null; else item.FESI = dr.GetDouble(17);
                if (dr.IsDBNull(18)) item.FEMN = null; else item.FEMN = dr.GetDouble(18);
                if (dr.IsDBNull(19)) item.FEP = null; else item.FEP = dr.GetDouble(19);
                if (dr.IsDBNull(20)) item.FES = null; else item.FES = dr.GetDouble(20);
                if (dr.IsDBNull(21)) item.FETI = null; else item.FETI = dr.GetDouble(21);
                if (dr.IsDBNull(22)) item.GFESI = null; else item.GFESI = dr.GetDouble(22);
                if (dr.IsDBNull(23)) item.GFES = null; else item.GFES = dr.GetDouble(23);
                if (dr.IsDBNull(24)) item.ZHASIO2 = null; else item.ZHASIO2 = dr.GetDouble(24);
                if (dr.IsDBNull(25)) item.ZHACAO = null; else item.ZHACAO = dr.GetDouble(25);
                if (dr.IsDBNull(26)) item.ZHAMGO = null; else item.ZHAMGO = dr.GetDouble(26);
                if (dr.IsDBNull(27)) item.ZHAAL2O3 = null; else item.ZHAAL2O3 = dr.GetDouble(27);
                if (dr.IsDBNull(28)) item.ZHAS = null; else item.ZHAS = dr.GetDouble(28);
                if (dr.IsDBNull(29)) item.ZHATIO2 = null; else item.ZHATIO2 = dr.GetDouble(29);
                if (dr.IsDBNull(30)) item.ZHAR2 = null; else item.ZHAR2 = dr.GetDouble(30);
                if (dr.IsDBNull(31)) item.ZHAR3 = null; else item.ZHAR3 = dr.GetDouble(31);
                if (dr.IsDBNull(32)) item.ZHAR4 = null; else item.ZHAR4 = dr.GetDouble(32);
                if (dr.IsDBNull(33)) item.ZHAR5 = null; else item.ZHAR5 = dr.GetDouble(33);
                if (dr.IsDBNull(34)) item.SHENDU = null; else item.SHENDU = dr.GetDouble(34);
                if (dr.IsDBNull(35)) item.JIAODU = null; else item.JIAODU = dr.GetDouble(35);
                if (dr.IsDBNull(36)) item.DANILIANG = null; else item.DANILIANG = dr.GetDouble(36);
                if (dr.IsDBNull(37)) item.ZHAYANG = ""; else item.ZHAYANG = dr.GetString(37);
                if (dr.IsDBNull(38)) item.SHUIWEN1 = null; else item.SHUIWEN1 = dr.GetDouble(38);
                if (dr.IsDBNull(39)) item.SHUIWEN2 = null; else item.SHUIWEN2 = dr.GetDouble(39);
                if (dr.IsDBNull(40)) item.PNQK1 = ""; else item.PNQK1 = dr.GetString(40);
                if (dr.IsDBNull(41)) item.PNQK2 = ""; else item.PNQK2 = dr.GetString(41);
                if (dr.IsDBNull(42)) item.CTC = ""; else item.CTC = dr.GetString(42);
                if (dr.IsDBNull(43)) item.LZQK = null; else item.LZQK = dr.GetDateTime(43);
                if (dr.IsDBNull(44)) item.BZ = ""; else item.BZ = dr.GetString(44);
                if (dr.IsDBNull(45)) item.BANCI = ""; else item.BANCI = dr.GetString(45);
                if (dr.IsDBNull(46)) item.BANLUCI = 0; else item.BANLUCI = dr.GetInt32(46);
                if (dr.IsDBNull(47)) item.BANZU = ""; else item.BANZU = dr.GetString(47);
                if (dr.IsDBNull(48)) item.GONGZHANG = ""; else item.GONGZHANG = dr.GetString(48);
                if (dr.IsDBNull(49)) item.RIXUHAO = 0; else item.RIXUHAO = dr.GetInt32(49);
                if (dr.IsDBNull(50)) item.GZSJ = null; else item.GZSJ = dr.GetDateTime(50);
                if (dr.IsDBNull(51)) item.GZJG = null; else item.GZJG = dr.GetInt32(51);
                if (dr.IsDBNull(52)) item.BBCL = null; else item.BBCL = dr.GetDouble(52);
                item.RowId = dr.GetString(53);

                item.EndInit();
                item.ClearDataState();
                this.Add(item);
                hideList.Add(item);
            }
            dr.Close();
            con.Close();
            beginAdjustData = false;
            if (rq < DateTime.Now.AddHours(8))
            {
                ChutiefanganTable chutiefangan = new ChutiefanganTable();
                chutiefangan.LoadByGlRq(gaolu, rq);
                foreach (var item in chutiefangan)
                {
                    Ddluci luci = new Ddluci();
                    luci.GAOLU = gaolu;
                    luci.ZDSJ = rq.AddMinutes(item.SJ);
                    luci.BANCI = item.BANCI;
                    luci.BANLUCI = item.BANCIHAO;
                    luci.RIXUHAO = item.XUHAO;
                    if (!this.Contains(luci))
                    {
                        Bd_Banzu bz = Bd_Banzu.GetBanZu(gaolu, rq, luci.BANCI);
                        if (bz != null)
                        {
                            luci.BANZU = bz.BANZU;
                            luci.GONGZHANG = bz.LEADER;
                        }
                        this.Add(luci);
                    }
                }
            }
            this.hideList.Sort(Compare);
            List<Ddluci> list = this.Items as List<Ddluci>;
            list.Sort(Compare);


            ListChangedEventArgs arg = new ListChangedEventArgs(ListChangedType.Reset, -1);
            this.OnListChanged(arg);
        }

        public void AddNewDDLuci(DateTime sj, string banci)
        {
            if (this.lastLoadGaolu != 0 && this.lastLoadRq != null)
            {
                Ddluci luci = new Ddluci();
                luci.GAOLU = this.lastLoadGaolu;
                luci.ZDSJ = new DateTime(this.lastLoadRq.Value.Year, this.lastLoadRq.Value.Month, this.lastLoadRq.Value.Day, sj.Hour, sj.Minute, 0);
                luci.BANCI = banci;
                if (!this.Contains(luci))
                {
                    Bd_Banzu bz = Bd_Banzu.GetBanZu(luci.GAOLU, lastLoadRq.Value, luci.BANCI);
                    if (bz != null)
                    {
                        luci.BANZU = bz.BANZU;
                        luci.GONGZHANG = bz.LEADER;
                    }
                    this.Add(luci);  
                    List<Ddluci> list = this.Items as List<Ddluci>;
                     list.Sort(Compare);
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("已有相同正点时间");
                }
            }
        }

        public void Load(string whereSql)
        {
     
            beginAdjustData = true;
            this.Clear();
            hideList.Clear();
            OracleConnection con = new OracleConnection(LTZN.Properties.Settings.Default.ltznConnectionString);
            con.Open();
            OracleCommand selectCmd = new OracleCommand();
            selectCmd.Connection = con;
            selectCmd.CommandText = "SELECT GAOLU,ZDSJ,LUCI,DGSJ,KKSJ,DKSJ,TZSJ,WDSJ,QUCHU,GFELIANG,LFELIANG,PIANCHA,FELIANG,LIAOPI,GGUANSHU,FEWENDU,FEC,FESI,FEMN,FEP,FES,FETI,GFESI,GFES,ZHASIO2,ZHACAO,ZHAMGO,ZHAAL2O3,ZHAS,ZHATIO2,ZHAR2,ZHAR3,ZHAR4,ZHAR5,SHENDU,JIAODU,DANILIANG,ZHAYANG,SHUIWEN1,SHUIWEN2,PNQK1,PNQK2,CTC,LZQK,BZ,BANCI,BANLUCI,BANZU,GONGZHANG,RIXUHAO,GZSJ,round((KKSJ-lag(DKSJ) over (order by ZDSJ))*1440,0) as GZJG,BBCL,ROWID FROM DDLUCI WHERE " + whereSql +" Order by GAOLU,zdsj";

            OracleDataReader dr = selectCmd.ExecuteReader();
            while (dr.Read())
            {
                Ddluci item = new Ddluci();
                item.BeginInit();
                if (dr.IsDBNull(0)) item.GAOLU = 0; else item.GAOLU = dr.GetInt32(0);
                if (dr.IsDBNull(1)) item.ZDSJ = DateTime.Now; else item.ZDSJ = dr.GetDateTime(1);
                if (dr.IsDBNull(2)) item.LUCI = ""; else item.LUCI = dr.GetString(2);
                if (dr.IsDBNull(3)) item.DGSJ = null; else item.DGSJ = dr.GetDateTime(3);
                if (dr.IsDBNull(4)) item.KKSJ = null; else item.KKSJ = dr.GetDateTime(4);
                if (dr.IsDBNull(5)) item.DKSJ = null; else item.DKSJ = dr.GetDateTime(5);
                if (dr.IsDBNull(6)) item.TZSJ = null; else item.TZSJ = dr.GetDateTime(6);
                if (dr.IsDBNull(7)) item.WDSJ = null; else item.WDSJ = dr.GetDouble(7);
                if (dr.IsDBNull(8)) item.QUCHU = ""; else item.QUCHU = dr.GetString(8);
                if (dr.IsDBNull(9)) item.GFELIANG = null; else item.GFELIANG = dr.GetDouble(9);
                if (dr.IsDBNull(10)) item.LFELIANG = null; else item.LFELIANG = dr.GetDouble(10);
                if (dr.IsDBNull(11)) item.PIANCHA = null; else item.PIANCHA = dr.GetDouble(11);
                if (dr.IsDBNull(12)) item.FELIANG = null; else item.FELIANG = dr.GetDouble(12);
                if (dr.IsDBNull(13)) item.LIAOPI = null; else item.LIAOPI = dr.GetDouble(13);
                if (dr.IsDBNull(14)) item.GGUANSHU = null; else item.GGUANSHU = dr.GetDouble(14);
                if (dr.IsDBNull(15)) item.FEWENDU = null; else item.FEWENDU = dr.GetDouble(15);
                if (dr.IsDBNull(16)) item.FEC = null; else item.FEC = dr.GetDouble(16);
                if (dr.IsDBNull(17)) item.FESI = null; else item.FESI = dr.GetDouble(17);
                if (dr.IsDBNull(18)) item.FEMN = null; else item.FEMN = dr.GetDouble(18);
                if (dr.IsDBNull(19)) item.FEP = null; else item.FEP = dr.GetDouble(19);
                if (dr.IsDBNull(20)) item.FES = null; else item.FES = dr.GetDouble(20);
                if (dr.IsDBNull(21)) item.FETI = null; else item.FETI = dr.GetDouble(21);
                if (dr.IsDBNull(22)) item.GFESI = null; else item.GFESI = dr.GetDouble(22);
                if (dr.IsDBNull(23)) item.GFES = null; else item.GFES = dr.GetDouble(23);
                if (dr.IsDBNull(24)) item.ZHASIO2 = null; else item.ZHASIO2 = dr.GetDouble(24);
                if (dr.IsDBNull(25)) item.ZHACAO = null; else item.ZHACAO = dr.GetDouble(25);
                if (dr.IsDBNull(26)) item.ZHAMGO = null; else item.ZHAMGO = dr.GetDouble(26);
                if (dr.IsDBNull(27)) item.ZHAAL2O3 = null; else item.ZHAAL2O3 = dr.GetDouble(27);
                if (dr.IsDBNull(28)) item.ZHAS = null; else item.ZHAS = dr.GetDouble(28);
                if (dr.IsDBNull(29)) item.ZHATIO2 = null; else item.ZHATIO2 = dr.GetDouble(29);
                if (dr.IsDBNull(30)) item.ZHAR2 = null; else item.ZHAR2 = dr.GetDouble(30);
                if (dr.IsDBNull(31)) item.ZHAR3 = null; else item.ZHAR3 = dr.GetDouble(31);
                if (dr.IsDBNull(32)) item.ZHAR4 = null; else item.ZHAR4 = dr.GetDouble(32);
                if (dr.IsDBNull(33)) item.ZHAR5 = null; else item.ZHAR5 = dr.GetDouble(33);
                if (dr.IsDBNull(34)) item.SHENDU = null; else item.SHENDU = dr.GetDouble(34);
                if (dr.IsDBNull(35)) item.JIAODU = null; else item.JIAODU = dr.GetDouble(35);
                if (dr.IsDBNull(36)) item.DANILIANG = null; else item.DANILIANG = dr.GetDouble(36);
                if (dr.IsDBNull(37)) item.ZHAYANG = ""; else item.ZHAYANG = dr.GetString(37);
                if (dr.IsDBNull(38)) item.SHUIWEN1 = null; else item.SHUIWEN1 = dr.GetDouble(38);
                if (dr.IsDBNull(39)) item.SHUIWEN2 = null; else item.SHUIWEN2 = dr.GetDouble(39);
                if (dr.IsDBNull(40)) item.PNQK1 = ""; else item.PNQK1 = dr.GetString(40);
                if (dr.IsDBNull(41)) item.PNQK2 = ""; else item.PNQK2 = dr.GetString(41);
                if (dr.IsDBNull(42)) item.CTC = ""; else item.CTC = dr.GetString(42);
                if (dr.IsDBNull(43)) item.LZQK = null; else item.LZQK = dr.GetDateTime(43);
                if (dr.IsDBNull(44)) item.BZ = ""; else item.BZ = dr.GetString(44);
                if (dr.IsDBNull(45)) item.BANCI = ""; else item.BANCI = dr.GetString(45);
                if (dr.IsDBNull(46)) item.BANLUCI = 0; else item.BANLUCI = dr.GetInt32(46);
                if (dr.IsDBNull(47)) item.BANZU = ""; else item.BANZU = dr.GetString(47);
                if (dr.IsDBNull(48)) item.GONGZHANG = ""; else item.GONGZHANG = dr.GetString(48);
                if (dr.IsDBNull(49)) item.RIXUHAO = 0; else item.RIXUHAO = dr.GetInt32(49);
                if (dr.IsDBNull(50)) item.GZSJ = null; else item.GZSJ = dr.GetDateTime(50);
                if (dr.IsDBNull(51)) item.GZJG = null; else item.GZJG = dr.GetInt32(51);
                if (dr.IsDBNull(52)) item.BBCL = null; else item.BBCL = dr.GetDouble(52);
                item.RowId = dr.GetString(53);

                item.EndInit();
                item.ClearDataState();
                this.Add(item);
                hideList.Add(item);
            }
            dr.Close();
            con.Close();
            beginAdjustData = false;
           
            ListChangedEventArgs arg = new ListChangedEventArgs(ListChangedType.Reset, -1);
            this.OnListChanged(arg);
        }

        public DateTime? GetMaxDKSJ(int gaolu, DateTime rq)
        {
            DateTime? result = null;
            OracleConnection con = new OracleConnection(LTZN.Properties.Settings.Default.ltznConnectionString);
            con.Open();
            OracleCommand selectCmd = new OracleCommand();
            selectCmd.Connection = con;
            selectCmd.CommandText = "select max(dksj) from ddluci where  gaolu=:GAOLU and zdsj<:RQ  and dksj is not null";
            selectCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = gaolu;
            selectCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = rq;
            OracleDataReader dr = selectCmd.ExecuteReader();
            if (dr.Read())
            {
                result = dr.IsDBNull(0) ? null : (DateTime?)dr.GetDateTime(0);
            }
            dr.Close();
            con.Close();
            return result;

        }

        public int GetLIAOPI(int gaolu, DateTime dksj1, DateTime dksj2)
        {
            int result = 0;
            OracleConnection con = new OracleConnection(LTZN.Properties.Settings.Default.ltznConnectionString);
            con.Open();
            OracleCommand selectCmd = new OracleCommand();
            selectCmd.Connection = con;
            selectCmd.CommandText = "select count(*) from rbliaoxian where gaolu=:GAOLU and t>:DKSJ1 and t<=:DKSJ2";
            selectCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = gaolu;
            selectCmd.Parameters.Add(":DKSJ1", OracleType.DateTime).Value = dksj1;
            selectCmd.Parameters.Add(":DKSJ2", OracleType.DateTime).Value = dksj2;
            OracleDataReader dr = selectCmd.ExecuteReader();
            if (dr.Read())
            {
                result = dr.IsDBNull(0) ? 0 : dr.GetInt32(0);
            }
            dr.Close();
            con.Close();
            return result;
        }

        private int Compare(Ddluci x, Ddluci y)
        {
            if (x.ZDSJ > y.ZDSJ)
                return 1;
            else if (x.ZDSJ == y.ZDSJ)
                return 0;
            else
                return -1;
        }

        private int lastEditGaolu = 0;  //堵口时间 最早的高炉
        private DateTime? lastEditZdsj = null; //


        public override void ExeOther(OracleConnection con)
        {
            con.Open();
            if (lastEditGaolu != 0 && lastEditZdsj != null)
            {
                OracleCommand dksjCmd = new OracleCommand();
                dksjCmd.Connection = con;
                dksjCmd.CommandText = "UPDATEDKSJ";
                dksjCmd.CommandType = CommandType.StoredProcedure;
                dksjCmd.Parameters.Add("P_GAOLU", OracleType.Int32).Value = lastEditGaolu;
                dksjCmd.Parameters.Add("P_ZDSJ", OracleType.DateTime).Value = lastEditZdsj.Value;
                dksjCmd.ExecuteNonQuery();
            }
            con.Close();
        }
        protected override void OnListChanged(ListChangedEventArgs e)
        {
            if (e.ListChangedType == ListChangedType.ItemChanged)
            {
                if (e.PropertyDescriptor.Name == "DKSJ")
                {
                    int curGAOLU = this[e.NewIndex].GAOLU;
                    DateTime curZDSJ = this[e.NewIndex].ZDSJ;

                    if (lastEditGaolu != curGAOLU || lastEditZdsj == null || lastEditZdsj > this[e.NewIndex].ZDSJ)
                    {
                        lastEditGaolu = curGAOLU;
                        lastEditZdsj = curZDSJ;
                    }
                }
            }
            base.OnListChanged(e);
        }
    }

    //出铁质量
    public partial class DdluciQa : DbEntity
    {

        private string _LUCI = "";
        private int _GAOLU;
        private System.DateTime _ZDSJ = DateTime.Now;
        private System.Nullable<double> _FEC;
        private System.Nullable<double> _FESI;
        private System.Nullable<double> _FEMN;
        private System.Nullable<double> _FEP;
        private System.Nullable<double> _FES;
        private System.Nullable<double> _FETI;

        private System.Nullable<double> _ZHASIO2;
        private System.Nullable<double> _ZHACAO;
        private System.Nullable<double> _ZHAMGO;
        private System.Nullable<double> _ZHAAL2O3;
        private System.Nullable<double> _ZHAS;
        private System.Nullable<double> _ZHATIO2;
        private System.Nullable<double> _ZHAR2;
        private string _BANCI = "";
        private int _BANLUCI;
        private System.Nullable<double> _ZHAR3;
        private System.Nullable<double> _ZHAR4;
        private System.Nullable<double> _ZHAR5;
        [DisplayName("高炉")]
        public int GAOLU
        {
            get
            {
                return this._GAOLU;
            }
            set
            {
                if (!_GAOLU.Equals(value))
                {
                    _GAOLU = value;
                    RaisePropertyChanged("GAOLU", true);
                }

            }
        }

        [DisplayName("正点时间")]
        public System.DateTime ZDSJ
        {
            get
            {
                return this._ZDSJ;
            }
            set
            {
                if (!_ZDSJ.Equals(value))
                {
                    _ZDSJ = value;
                    RaisePropertyChanged("ZDSJ", true);
                }

            }
        }

        [DisplayName("炉次")]
        public string JYLUCI
        {
            get
            {
                return GAOLU.ToString() + RIXUHAO.ToString("00");
            }
        }

        private int _RIXUHAO;
        [DisplayName("日序号")]
        public int RIXUHAO
        {
            get { return _RIXUHAO; }
            set
            {
                if (_RIXUHAO != value)
                {
                    _RIXUHAO = value;
                    RaisePropertyChanged("RIXUHAO");
                }
            }
        }

        [DisplayName("班次")]
        public string BANCI
        {
            get
            {
                return this._BANCI;
            }
            set
            {
                if (!_BANCI.Equals(value))
                {
                    _BANCI = value;
                    RaisePropertyChanged("BANCI", true);
                }

            }
        }

        [DisplayName("班次号")]
        public int BANLUCI
        {
            get
            {
                return this._BANLUCI;
            }
            set
            {
                if (!_BANLUCI.Equals(value))
                {
                    _BANLUCI = value;
                    RaisePropertyChanged("BANLUCI", true);
                }

            }
        }

        private string _BANZU = "";
        [DisplayName("班组")]
        public string BANZU
        {
            get { return _BANZU; }
            set
            {
                if (_BANZU != value)
                {
                    _BANZU = value;
                    RaisePropertyChanged("BANZU");
                }
            }
        }

        private string _GONGZHANG = "";
        [DisplayName("工长")]
        public string GONGZHANG
        {
            get { return _GONGZHANG; }
            set
            {
                if (_GONGZHANG != value)
                {
                    _GONGZHANG = value;
                    RaisePropertyChanged("GONGZHANG");
                }
            }
        }

        [DisplayName("炉次")]
        public string LUCI
        {
            get
            {
                return this._LUCI;
            }
            set
            {
                if (!_LUCI.Equals(value))
                {
                    _LUCI = value;
                    RaisePropertyChanged("LUCI", true);
                }

            }
        }

        [DisplayName("铁水C")]
        public System.Nullable<double> FEC
        {
            get
            {
                return this._FEC;
            }
            set
            {
                if (!_FEC.Equals(value))
                {
                    _FEC = value;
                    RaisePropertyChanged("FEC", true);
                }

            }
        }

        [DisplayName("铁水Si")]
        public System.Nullable<double> FESI
        {
            get
            {
                return this._FESI;
            }
            set
            {
                if (!_FESI.Equals(value))
                {
                    _FESI = value;
                    RaisePropertyChanged("FESI", true);
                }

            }
        }

        [DisplayName("铁水Mn")]
        public System.Nullable<double> FEMN
        {
            get
            {
                return this._FEMN;
            }
            set
            {
                if (!_FEMN.Equals(value))
                {
                    _FEMN = value;
                    RaisePropertyChanged("FEMN", true);
                }

            }
        }

        [DisplayName("铁水P")]
        public System.Nullable<double> FEP
        {
            get
            {
                return this._FEP;
            }
            set
            {
                if (!_FEP.Equals(value))
                {
                    _FEP = value;
                    RaisePropertyChanged("FEP", true);
                }

            }
        }

        [DisplayName("铁水S")]
        public System.Nullable<double> FES
        {
            get
            {
                return this._FES;
            }
            set
            {
                if (!_FES.Equals(value))
                {
                    _FES = value;
                    RaisePropertyChanged("FES", true);
                }

            }
        }

        [DisplayName("铁水Ti")]
        public System.Nullable<double> FETI
        {
            get
            {
                return this._FETI;
            }
            set
            {
                if (!_FETI.Equals(value))
                {
                    _FETI = value;
                    RaisePropertyChanged("FETI", true);
                }

            }
        }

        [DisplayName("渣SiO2")]
        public System.Nullable<double> ZHASIO2
        {
            get
            {
                return this._ZHASIO2;
            }
            set
            {
                if (!_ZHASIO2.Equals(value))
                {
                    _ZHASIO2 = value;
                    RaisePropertyChanged("ZHASIO2", true);
                    RaisePropertyChanged("ZHAR2");
                    RaisePropertyChanged("ZHAR3");
                    RaisePropertyChanged("ZHAR4");
                    RaisePropertyChanged("ZHAR5");
                }

            }
        }

        [DisplayName("渣CaO")]
        public System.Nullable<double> ZHACAO
        {
            get
            {
                return this._ZHACAO;
            }
            set
            {
                if (!_ZHACAO.Equals(value))
                {
                    _ZHACAO = value;
                    RaisePropertyChanged("ZHACAO", true);
                    RaisePropertyChanged("ZHAR2");
                    RaisePropertyChanged("ZHAR3");
                    RaisePropertyChanged("ZHAR4");
                    RaisePropertyChanged("ZHAR5");
                }

            }
        }

        [DisplayName("渣MgO")]
        public System.Nullable<double> ZHAMGO
        {
            get
            {
                return this._ZHAMGO;
            }
            set
            {
                if (!_ZHAMGO.Equals(value))
                {
                    _ZHAMGO = value;
                    RaisePropertyChanged("ZHAMGO", true);
                    RaisePropertyChanged("ZHAR2");
                    RaisePropertyChanged("ZHAR3");
                    RaisePropertyChanged("ZHAR4");
                    RaisePropertyChanged("ZHAR5");
                }

            }
        }

        [DisplayName("渣Al2O3")]
        public System.Nullable<double> ZHAAL2O3
        {
            get
            {
                return this._ZHAAL2O3;
            }
            set
            {
                if (!_ZHAAL2O3.Equals(value))
                {
                    _ZHAAL2O3 = value;
                    RaisePropertyChanged("ZHAAL2O3", true);
                    RaisePropertyChanged("ZHAR2");
                    RaisePropertyChanged("ZHAR3");
                    RaisePropertyChanged("ZHAR4");
                    RaisePropertyChanged("ZHAR5");
                }

            }
        }

        [DisplayName("渣S")]
        public System.Nullable<double> ZHAS
        {
            get
            {
                return this._ZHAS;
            }
            set
            {
                if (!_ZHAS.Equals(value))
                {
                    _ZHAS = value;
                    RaisePropertyChanged("ZHAS", true);
                }

            }
        }

        [DisplayName("渣TiO2")]
        public System.Nullable<double> ZHATIO2
        {
            get
            {
                return this._ZHATIO2;
            }
            set
            {
                if (!_ZHATIO2.Equals(value))
                {
                    _ZHATIO2 = value;
                    RaisePropertyChanged("ZHATIO2", true);
                }

            }
        }

        [DisplayName("渣R2")]
        public System.Nullable<double> ZHAR2
        {
            get
            {
                _ZHAR2 = this._ZHACAO / this._ZHASIO2;
                if (_ZHAR2 == double.NegativeInfinity)
                    _ZHAR2 = null;
                return _ZHAR2;
            }
            //set
            //{
            //    if (!_ZHAR2.Equals(value))
            //    {
            //        _ZHAR2 = value;
            //        RaisePropertyChanged("ZHAR2", true);
            //    }

            //}
        }

        [DisplayName("渣R3")]
        public System.Nullable<double> ZHAR3
        {
            get
            {
                _ZHAR3 = (this._ZHACAO + this._ZHAMGO) / this._ZHASIO2;
                if (_ZHAR3 == double.NegativeInfinity)
                    _ZHAR3 = null;
                return _ZHAR3;
            }
            //set
            //{
            //    if (!_ZHAR3.Equals(value))
            //    {
            //        _ZHAR3 = value;
            //        RaisePropertyChanged("ZHAR3", true);
            //    }

            //}
        }

        [DisplayName("渣R4")]
        public System.Nullable<double> ZHAR4
        {
            get
            {
                _ZHAR4 = (this._ZHACAO + this._ZHAMGO) / (this._ZHASIO2 + this._ZHAAL2O3);
                if (_ZHAR4 == double.NegativeInfinity)
                    _ZHAR4 = null;
                return _ZHAR4;
            }
            //set
            //{
            //    if (!_ZHAR4.Equals(value))
            //    {
            //        _ZHAR4 = value;
            //        RaisePropertyChanged("ZHAR4", true);
            //    }

            //}
        }

        [DisplayName("渣MgO/Al2O3")]
        public System.Nullable<double> ZHAR5
        {
            get
            {
                _ZHAR5 = this._ZHAMGO / this._ZHAAL2O3;
                if (_ZHAR5 == double.NegativeInfinity)
                    _ZHAR5 = null;
                return _ZHAR5;
            }
            //set
            //{
            //    if (!_ZHAR5.Equals(value))
            //    {
            //        _ZHAR5 = value;
            //        RaisePropertyChanged("ZHAR5", true);
            //    }

            //}
        }


        public bool CheckPk(OracleTransaction trans)
        {
            OracleConnection con = trans.Connection;
            OracleCommand checkCmd = new OracleCommand();
            checkCmd.Connection = con;
            checkCmd.Transaction = trans;
            checkCmd.CommandText = "SELECT ROWID FROM DDLUCI WHERE GAOLU=:GAOLU AND ZDSJ=:ZDSJ";
            checkCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = this.GAOLU;
            checkCmd.Parameters.Add(":ZDSJ", OracleType.DateTime).Value = this.ZDSJ;
            OracleDataReader dr = checkCmd.ExecuteReader();
            bool result = false;
            if (dr.Read())
            {
                OracleString rowid = dr.GetOracleString(0);
                if (!rowid.IsNull)
                {
                    this.RowId = rowid;
                    result = true;
                }
            }
            dr.Close();
            return result;
        }

        public override void DbInsert(OracleTransaction trans)
        {
            //InsertSql,InsertPara,AutoSql
            if (DataState == DataStateType.Add)
            {
                OracleConnection con = trans.Connection;
                if (CheckPk(trans))
                {
                    OracleCommand updateCmd = new OracleCommand();
                    updateCmd.Connection = con;
                    updateCmd.Transaction = trans;
                    updateCmd.CommandText = "UPDATE DDLUCI SET FEC=:FEC,FESI=:FESI,FEMN=:FEMN,FEP=:FEP,FES=:FES,FETI=:FETI,ZHASIO2=:ZHASIO2,ZHACAO=:ZHACAO,ZHAMGO=:ZHAMGO,ZHAAL2O3=:ZHAAL2O3,ZHAS=:ZHAS,ZHATIO2=:ZHATIO2,ZHAR2=:ZHAR2,ZHAR3=:ZHAR3,ZHAR4=:ZHAR4,ZHAR5=:ZHAR5 WHERE GAOLU=:GAOLU AND ZDSJ=:ZDSJ";
                    updateCmd.Parameters.Add(":FEC", OracleType.Double, 22).Value = (object)this.FEC ?? DBNull.Value;
                    updateCmd.Parameters.Add(":FESI", OracleType.Double, 22).Value = (object)this.FESI ?? DBNull.Value;
                    updateCmd.Parameters.Add(":FEMN", OracleType.Double, 22).Value = (object)this.FEMN ?? DBNull.Value;
                    updateCmd.Parameters.Add(":FEP", OracleType.Double, 22).Value = (object)this.FEP ?? DBNull.Value;
                    updateCmd.Parameters.Add(":FES", OracleType.Double, 22).Value = (object)this.FES ?? DBNull.Value;
                    updateCmd.Parameters.Add(":FETI", OracleType.Double, 22).Value = (object)this.FETI ?? DBNull.Value;
                    updateCmd.Parameters.Add(":ZHASIO2", OracleType.Double, 22).Value = (object)this.ZHASIO2 ?? DBNull.Value;
                    updateCmd.Parameters.Add(":ZHACAO", OracleType.Double, 22).Value = (object)this.ZHACAO ?? DBNull.Value;
                    updateCmd.Parameters.Add(":ZHAMGO", OracleType.Double, 22).Value = (object)this.ZHAMGO ?? DBNull.Value;
                    updateCmd.Parameters.Add(":ZHAAL2O3", OracleType.Double, 22).Value = (object)this.ZHAAL2O3 ?? DBNull.Value;
                    updateCmd.Parameters.Add(":ZHAS", OracleType.Double, 22).Value = (object)this.ZHAS ?? DBNull.Value;
                    updateCmd.Parameters.Add(":ZHATIO2", OracleType.Double, 22).Value = (object)this.ZHATIO2 ?? DBNull.Value;
                    updateCmd.Parameters.Add(":ZHAR2", OracleType.Double, 22).Value = (object)this.ZHAR2 ?? DBNull.Value;
                    updateCmd.Parameters.Add(":ZHAR3", OracleType.Double, 22).Value = (object)this.ZHAR3 ?? DBNull.Value;
                    updateCmd.Parameters.Add(":ZHAR4", OracleType.Double, 22).Value = (object)this.ZHAR4 ?? DBNull.Value;
                    updateCmd.Parameters.Add(":ZHAR5", OracleType.Double, 22).Value = (object)this.ZHAR5 ?? DBNull.Value;
                    updateCmd.Parameters.Add(":GAOLU", OracleType.Int32, 22).Value = this.GAOLU;
                    updateCmd.Parameters.Add(":ZDSJ", OracleType.DateTime, 7).Value = this.ZDSJ;
                    updateCmd.ExecuteOracleNonQuery(out RowId);
                    return;
                }
                OracleCommand insertCmd = new OracleCommand();
                insertCmd.Connection = con;
                insertCmd.Transaction = trans;
                insertCmd.CommandText = "INSERT INTO DDLUCI(GAOLU,ZDSJ,LUCI,FEC,FESI,FEMN,FEP,FES,FETI,ZHASIO2,ZHACAO,ZHAMGO,ZHAAL2O3,ZHAS,ZHATIO2,ZHAR2,ZHAR3,ZHAR4,ZHAR5,BANCI,BANLUCI,BANZU,GONGZHANG,RIXUHAO) VALUES(:GAOLU,:ZDSJ,:LUCI,:FEC,:FESI,:FEMN,:FEP,:FES,:FETI,:ZHASIO2,:ZHACAO,:ZHAMGO,:ZHAAL2O3,:ZHAS,:ZHATIO2,:ZHAR2,:ZHAR3,:ZHAR4,:ZHAR5,:BANCI,:BANLUCI,:BANZU,:GONGZHANG,:RIXUHAO)";
                insertCmd.Parameters.Add(":GAOLU", OracleType.Int32, 22).Value = this.GAOLU;
                insertCmd.Parameters.Add(":ZDSJ", OracleType.DateTime, 7).Value = this.ZDSJ;
                insertCmd.Parameters.Add(":LUCI", OracleType.VarChar, 20).Value = this.LUCI;
                insertCmd.Parameters.Add(":FEC", OracleType.Double, 22).Value = (object)this.FEC ?? DBNull.Value;
                insertCmd.Parameters.Add(":FESI", OracleType.Double, 22).Value = (object)this.FESI ?? DBNull.Value;
                insertCmd.Parameters.Add(":FEMN", OracleType.Double, 22).Value = (object)this.FEMN ?? DBNull.Value;
                insertCmd.Parameters.Add(":FEP", OracleType.Double, 22).Value = (object)this.FEP ?? DBNull.Value;
                insertCmd.Parameters.Add(":FES", OracleType.Double, 22).Value = (object)this.FES ?? DBNull.Value;
                insertCmd.Parameters.Add(":FETI", OracleType.Double, 22).Value = (object)this.FETI ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHASIO2", OracleType.Double, 22).Value = (object)this.ZHASIO2 ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHACAO", OracleType.Double, 22).Value = (object)this.ZHACAO ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHAMGO", OracleType.Double, 22).Value = (object)this.ZHAMGO ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHAAL2O3", OracleType.Double, 22).Value = (object)this.ZHAAL2O3 ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHAS", OracleType.Double, 22).Value = (object)this.ZHAS ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHATIO2", OracleType.Double, 22).Value = (object)this.ZHATIO2 ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHAR2", OracleType.Double, 22).Value = (object)this.ZHAR2 ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHAR3", OracleType.Double, 22).Value = (object)this.ZHAR3 ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHAR4", OracleType.Double, 22).Value = (object)this.ZHAR4 ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHAR5", OracleType.Double, 22).Value = (object)this.ZHAR5 ?? DBNull.Value;
                insertCmd.Parameters.Add(":BANCI", OracleType.VarChar, 20).Value = this.BANCI;
                insertCmd.Parameters.Add(":BANLUCI", OracleType.Int32, 22).Value = this.BANLUCI;
                insertCmd.Parameters.Add(":BANZU", OracleType.VarChar, 50).Value = this.BANZU;
                insertCmd.Parameters.Add(":GONGZHANG", OracleType.VarChar, 50).Value = this.GONGZHANG;
                insertCmd.Parameters.Add(":RIXUHAO", OracleType.Int32).Value = this.RIXUHAO;
                insertCmd.ExecuteOracleNonQuery(out RowId);
            }



        }
        public override void DbDelete(OracleTransaction trans)
        {
            //DelSql,DelPara
            if (DataState == DataStateType.Delete)
            {
                OracleConnection con = trans.Connection;
                OracleCommand deleteCmd = new OracleCommand();
                deleteCmd.Connection = con;
                deleteCmd.Transaction = trans;
                deleteCmd.CommandText = "DELETE FROM DDLUCI WHERE ROWID='" + RowId.Value + "'";

                deleteCmd.ExecuteNonQuery();
            }

        }
        public override void DbUpdate(OracleTransaction trans)
        {
            //UpdateSql,UpdatePara
            if (DataState == DataStateType.Update)
            {
                OracleConnection con = trans.Connection;
                OracleCommand updateCmd = new OracleCommand();
                updateCmd.Connection = con;
                updateCmd.Transaction = trans;
                updateCmd.CommandText = "UPDATE DDLUCI SET FEC=:FEC,FESI=:FESI,FEMN=:FEMN,FEP=:FEP,FES=:FES,FETI=:FETI,ZHASIO2=:ZHASIO2,ZHACAO=:ZHACAO,ZHAMGO=:ZHAMGO,ZHAAL2O3=:ZHAAL2O3,ZHAS=:ZHAS,ZHATIO2=:ZHATIO2,ZHAR2=:ZHAR2,ZHAR3=:ZHAR3,ZHAR4=:ZHAR4,ZHAR5=:ZHAR5 WHERE ROWID='" + RowId.Value + "'";
                updateCmd.Parameters.Add(":FEC", OracleType.Double, 22).Value = (object)this.FEC ?? DBNull.Value;
                updateCmd.Parameters.Add(":FESI", OracleType.Double, 22).Value = (object)this.FESI ?? DBNull.Value;
                updateCmd.Parameters.Add(":FEMN", OracleType.Double, 22).Value = (object)this.FEMN ?? DBNull.Value;
                updateCmd.Parameters.Add(":FEP", OracleType.Double, 22).Value = (object)this.FEP ?? DBNull.Value;
                updateCmd.Parameters.Add(":FES", OracleType.Double, 22).Value = (object)this.FES ?? DBNull.Value;
                updateCmd.Parameters.Add(":FETI", OracleType.Double, 22).Value = (object)this.FETI ?? DBNull.Value;
                updateCmd.Parameters.Add(":ZHASIO2", OracleType.Double, 22).Value = (object)this.ZHASIO2 ?? DBNull.Value;
                updateCmd.Parameters.Add(":ZHACAO", OracleType.Double, 22).Value = (object)this.ZHACAO ?? DBNull.Value;
                updateCmd.Parameters.Add(":ZHAMGO", OracleType.Double, 22).Value = (object)this.ZHAMGO ?? DBNull.Value;
                updateCmd.Parameters.Add(":ZHAAL2O3", OracleType.Double, 22).Value = (object)this.ZHAAL2O3 ?? DBNull.Value;
                updateCmd.Parameters.Add(":ZHAS", OracleType.Double, 22).Value = (object)this.ZHAS ?? DBNull.Value;
                updateCmd.Parameters.Add(":ZHATIO2", OracleType.Double, 22).Value = (object)this.ZHATIO2 ?? DBNull.Value;
                updateCmd.Parameters.Add(":ZHAR2", OracleType.Double, 22).Value = (object)this.ZHAR2 ?? DBNull.Value;
                updateCmd.Parameters.Add(":ZHAR3", OracleType.Double, 22).Value = (object)this.ZHAR3 ?? DBNull.Value;
                updateCmd.Parameters.Add(":ZHAR4", OracleType.Double, 22).Value = (object)this.ZHAR4 ?? DBNull.Value;
                updateCmd.Parameters.Add(":ZHAR5", OracleType.Double, 22).Value = (object)this.ZHAR5 ?? DBNull.Value;
                updateCmd.ExecuteNonQuery();
            }
        }

        public override bool Equals(object obj)
        {
            DdluciQa other = obj as DdluciQa;
            if (other != null)
            {
                if (this.GAOLU == other.GAOLU && this.ZDSJ == other.ZDSJ)
                    return true;
            }
            return false;
        }

        public override int GetHashCode()
        {
            // int hash = this.ZDSJ.Year * 10000 + this.ZDSJ.Month * 10000 + this.ZDSJ.Day * 1000 + this.ZDSJ.Hour * 100 + this.ZDSJ.Minute + GAOLU;
            return base.GetHashCode();
        }
    }
    public partial class DdluciQaTable : DbEntityTable<DdluciQa>
    {
        public void LoadByRQXuHao(DateTime riqi, string banci, int banluci)
        {
            beginAdjustData = true;
            this.Clear();
            hideList.Clear();
            OracleConnection con = new OracleConnection(LTZN.Properties.Settings.Default.ltznConnectionString);
            con.Open();
            OracleCommand selectCmd = new OracleCommand();
            selectCmd.Connection = con;
            selectCmd.CommandText = "SELECT GAOLU,ZDSJ,LUCI,FEC,FESI,FEMN,FEP,FES,FETI,ZHASIO2,ZHACAO,ZHAMGO,ZHAAL2O3,ZHAS,ZHATIO2,ZHAR2,ZHAR3,ZHAR4,ZHAR5,BANCI,BANLUCI,BANZU,GONGZHANG,RIXUHAO,ROWID FROM DDLUCI WHERE TRUNC(ZDSJ)=:RQ AND BANCI=:BANCI AND BANLUCI=:BANLUCI";
            selectCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = riqi;
            selectCmd.Parameters.Add(":BANCI", OracleType.VarChar, 50).Value = banci;
            selectCmd.Parameters.Add(":BANLUCI", OracleType.Int32).Value = banluci;
            OracleDataReader dr = selectCmd.ExecuteReader();
            while (dr.Read())
            {
                DdluciQa item = new DdluciQa();
                item.BeginInit();
                if (dr.IsDBNull(0)) item.GAOLU = 0; else item.GAOLU = dr.GetInt32(0);
                if (dr.IsDBNull(1)) item.ZDSJ = DateTime.Now; else item.ZDSJ = dr.GetDateTime(1);
                if (dr.IsDBNull(2)) item.LUCI = ""; else item.LUCI = dr.GetString(2);
                if (dr.IsDBNull(3)) item.FEC = null; else item.FEC = dr.GetDouble(3);
                if (dr.IsDBNull(4)) item.FESI = null; else item.FESI = dr.GetDouble(4);
                if (dr.IsDBNull(5)) item.FEMN = null; else item.FEMN = dr.GetDouble(5);
                if (dr.IsDBNull(6)) item.FEP = null; else item.FEP = dr.GetDouble(6);
                if (dr.IsDBNull(7)) item.FES = null; else item.FES = dr.GetDouble(7);
                if (dr.IsDBNull(8)) item.FETI = null; else item.FETI = dr.GetDouble(8);
                if (dr.IsDBNull(9)) item.ZHASIO2 = null; else item.ZHASIO2 = dr.GetDouble(9);
                if (dr.IsDBNull(10)) item.ZHACAO = null; else item.ZHACAO = dr.GetDouble(10);
                if (dr.IsDBNull(11)) item.ZHAMGO = null; else item.ZHAMGO = dr.GetDouble(11);
                if (dr.IsDBNull(12)) item.ZHAAL2O3 = null; else item.ZHAAL2O3 = dr.GetDouble(12);
                if (dr.IsDBNull(13)) item.ZHAS = null; else item.ZHAS = dr.GetDouble(13);
                if (dr.IsDBNull(14)) item.ZHATIO2 = null; else item.ZHATIO2 = dr.GetDouble(14);
                //if (dr.IsDBNull(15)) item.ZHAR2 = null; else item.ZHAR2 = dr.GetDouble(15);
                //if (dr.IsDBNull(16)) item.ZHAR3 = null; else item.ZHAR3 = dr.GetDouble(16);
                //if (dr.IsDBNull(17)) item.ZHAR4 = null; else item.ZHAR4 = dr.GetDouble(17);
                //if (dr.IsDBNull(18)) item.ZHAR5 = null; else item.ZHAR5 = dr.GetDouble(18);
                if (dr.IsDBNull(19)) item.BANCI = ""; else item.BANCI = dr.GetString(19);
                if (dr.IsDBNull(20)) item.BANLUCI = 0; else item.BANLUCI = dr.GetInt32(20);
                if (dr.IsDBNull(21)) item.BANZU = ""; else item.BANZU = dr.GetString(21);
                if (dr.IsDBNull(22)) item.GONGZHANG = ""; else item.GONGZHANG = dr.GetString(22);
                if (dr.IsDBNull(23)) item.RIXUHAO = 0; else item.RIXUHAO = dr.GetInt32(23);
                item.RowId = dr.GetString(24);

                item.EndInit();
                item.ClearDataState();
                this.Add(item);
                hideList.Add(item);
            }
            dr.Close();
            con.Close();
            beginAdjustData = false;
            if (riqi < DateTime.Now.AddHours(8))
            {
                Bd_GongxuTable gx = new Bd_GongxuTable();
                gx.LoadGaolu();
                foreach (var gaolu in gx)
                {
                    ChutiefanganTable chutiefangan = new ChutiefanganTable();
                    chutiefangan.LoadByGlRq(gaolu.ID, riqi);
                    foreach (var item in chutiefangan)
                    {
                        if (item.BANCI == banci && item.BANCIHAO == banluci)
                        {
                            DdluciQa luci = new DdluciQa();
                            luci.GAOLU = gaolu.ID;
                            luci.ZDSJ = riqi.AddMinutes(item.SJ);
                            luci.BANCI = item.BANCI;
                            luci.BANLUCI = item.BANCIHAO;
                            luci.RIXUHAO = item.XUHAO;
                            if (!this.Contains(luci))
                            {
                                Bd_Banzu bz = Bd_Banzu.GetBanZu(luci.GAOLU, riqi, luci.BANCI);
                                if (bz != null)
                                {
                                    luci.BANZU = bz.BANZU;
                                    luci.GONGZHANG = bz.LEADER;
                                }
                                this.Add(luci);

                            }
                        }
                    }
                }
            }
            ListChangedEventArgs arg = new ListChangedEventArgs(ListChangedType.Reset, -1);
            this.OnListChanged(arg);
        }

        public void LoadByGLRQ(int gaolu, DateTime rq)
        {
            beginAdjustData = true;
            this.Clear();
            hideList.Clear();
            OracleConnection con = new OracleConnection(LTZN.Properties.Settings.Default.ltznConnectionString);
            con.Open();
            OracleCommand selectCmd = new OracleCommand();
            selectCmd.Connection = con;
            selectCmd.CommandText = "SELECT GAOLU,ZDSJ,LUCI,FEC,FESI,FEMN,FEP,FES,FETI,ZHASIO2,ZHACAO,ZHAMGO,ZHAAL2O3,ZHAS,ZHATIO2,ZHAR2,ZHAR3,ZHAR4,ZHAR5,BANCI,BANLUCI,BANZU,GONGZHANG,RIXUHAO,ROWID FROM DDLUCI WHERE GAOLU=:GAOLU AND TRUNC(ZDSJ)=:RQ";
            selectCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = gaolu;
            selectCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = rq;

            OracleDataReader dr = selectCmd.ExecuteReader();
            while (dr.Read())
            {
                DdluciQa item = new DdluciQa();
                item.BeginInit();
                if (dr.IsDBNull(0)) item.GAOLU = 0; else item.GAOLU = dr.GetInt32(0);
                if (dr.IsDBNull(1)) item.ZDSJ = DateTime.Now; else item.ZDSJ = dr.GetDateTime(1);
                if (dr.IsDBNull(2)) item.LUCI = ""; else item.LUCI = dr.GetString(2);
                if (dr.IsDBNull(3)) item.FEC = null; else item.FEC = dr.GetDouble(3);
                if (dr.IsDBNull(4)) item.FESI = null; else item.FESI = dr.GetDouble(4);
                if (dr.IsDBNull(5)) item.FEMN = null; else item.FEMN = dr.GetDouble(5);
                if (dr.IsDBNull(6)) item.FEP = null; else item.FEP = dr.GetDouble(6);
                if (dr.IsDBNull(7)) item.FES = null; else item.FES = dr.GetDouble(7);
                if (dr.IsDBNull(8)) item.FETI = null; else item.FETI = dr.GetDouble(8);
                if (dr.IsDBNull(9)) item.ZHASIO2 = null; else item.ZHASIO2 = dr.GetDouble(9);
                if (dr.IsDBNull(10)) item.ZHACAO = null; else item.ZHACAO = dr.GetDouble(10);
                if (dr.IsDBNull(11)) item.ZHAMGO = null; else item.ZHAMGO = dr.GetDouble(11);
                if (dr.IsDBNull(12)) item.ZHAAL2O3 = null; else item.ZHAAL2O3 = dr.GetDouble(12);
                if (dr.IsDBNull(13)) item.ZHAS = null; else item.ZHAS = dr.GetDouble(13);
                if (dr.IsDBNull(14)) item.ZHATIO2 = null; else item.ZHATIO2 = dr.GetDouble(14);
                //if (dr.IsDBNull(15)) item.ZHAR2 = null; else item.ZHAR2 = dr.GetDouble(15);
                //if (dr.IsDBNull(16)) item.ZHAR3 = null; else item.ZHAR3 = dr.GetDouble(16);
                //if (dr.IsDBNull(17)) item.ZHAR4 = null; else item.ZHAR4 = dr.GetDouble(17);
                //if (dr.IsDBNull(18)) item.ZHAR5 = null; else item.ZHAR5 = dr.GetDouble(18);
                if (dr.IsDBNull(19)) item.BANCI = ""; else item.BANCI = dr.GetString(19);
                if (dr.IsDBNull(20)) item.BANLUCI = 0; else item.BANLUCI = dr.GetInt32(20);
                if (dr.IsDBNull(21)) item.BANZU = ""; else item.BANZU = dr.GetString(21);
                if (dr.IsDBNull(22)) item.GONGZHANG = ""; else item.GONGZHANG = dr.GetString(22);
                if (dr.IsDBNull(23)) item.RIXUHAO = 0; else item.RIXUHAO = dr.GetInt32(23);
                item.RowId = dr.GetString(24);

                item.EndInit();
                item.ClearDataState();
                this.Add(item);
                hideList.Add(item);
            }
            dr.Close();
            con.Close();
            beginAdjustData = false;
            if (rq < DateTime.Now.AddHours(8))
            {
                ChutiefanganTable chutiefangan = new ChutiefanganTable();
                chutiefangan.LoadByGlRq(gaolu, rq);
                foreach (var item in chutiefangan)
                {
                    DdluciQa luci = new DdluciQa();
                    luci.GAOLU = gaolu;
                    luci.ZDSJ = rq.AddMinutes(item.SJ);
                    luci.BANCI = item.BANCI;
                    luci.BANLUCI = item.BANCIHAO;
                    luci.RIXUHAO = item.XUHAO;
                    if (!this.Contains(luci))
                    {
                        Bd_Banzu bz = Bd_Banzu.GetBanZu(gaolu, rq, luci.BANCI);
                        if (bz != null)
                        {
                            luci.BANZU = bz.BANZU;
                            luci.GONGZHANG = bz.LEADER;
                        }
                        this.Add(luci);
                    }
                }
            }
            this.hideList.Sort(Compare);
            List<DdluciQa> list = this.Items as List<DdluciQa>;
            list.Sort(Compare);
            ListChangedEventArgs arg = new ListChangedEventArgs(ListChangedType.Reset, -1);
            this.OnListChanged(arg);

        }
        private int Compare(DdluciQa x, DdluciQa y)
        {
            if (x.ZDSJ > y.ZDSJ)
                return 1;
            else if (x.ZDSJ == y.ZDSJ)
                return 0;
            else
                return -1;
        }

    }
}
