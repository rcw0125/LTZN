using System;
using System.Data;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Reflection;
using System.Data.OracleClient;
using LTZN;

namespace ZHCDB
{
    public partial class Rbzlbd : DbEntity
    {
        private int? _GAOLU;
        private System.Nullable<System.DateTime> _SJ = null;
        private string _BC = "";
        private System.Nullable<double> _PICI;
        private System.Nullable<double> _JS;
        private System.Nullable<double> _SQ;
        private System.Nullable<double> _QT;
        private System.Nullable<double> _JT;
        private System.Nullable<double> _JD;
        private System.Nullable<double> _KP;
        private System.Nullable<double> _JP;
        private System.Nullable<double> _JTFH;
        private System.Nullable<double> _PML;
        private System.Nullable<double> _ZHFH;
        private System.Nullable<double> _ZL;
        private System.Nullable<double> _R2;
        private System.Nullable<double> _LLJD;
        private string _ZLSX = "";
        private System.Nullable<int> _LX;
        private System.Nullable<System.DateTime> _BLSJ = null;
        private string _BLYY = "";
        private string _CLCH = "";
        private System.Nullable<double> _LLTL;
        private string _BZ = "";
        private System.Nullable<double> _YQ;
        private System.Nullable<double> _LLBC;
        private string _BLJD = "";
        private System.Nullable<double> _LLJB;
        private System.Nullable<double> _SJKPB;
        private System.Nullable<double> _SLQPB;
        private System.Nullable<double> _KKPB;
        private System.Nullable<double> _KK;
        private System.Nullable<double> _QRPB;
        //ConStr 
        private System.Nullable<double> _RLPW;    //入炉品位
        private System.Nullable<double> _BY1;     //备用1
        private System.Nullable<double> _BY2;     //备用2

        [Query]
        [DisplayName("高炉")]
        public int? GAOLU
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
        [DisplayName("时间")]
        public System.Nullable<System.DateTime> SJ
        {
            get
            {
                return this._SJ;
            }
            set
            {
                if (!_SJ.Equals(value))
                {
                    _SJ = value;
                    RaisePropertyChanged("SJ", true);
                }

            }
        }
        
        [Query]
        [DisplayName("班次")]
        public string BC
        {
            get
            {
                return this._BC;
            }
            set
            {
                if (!_BC.Equals(value))
                {
                    _BC = value;
                    RaisePropertyChanged("BC", true);
                }

            }
        }
       
        [Query]
        [DisplayName("批次")]
        public System.Nullable<double> PICI
        {
            get
            {
                return this._PICI;
            }
            set
            {
                if (!_PICI.Equals(value))
                {
                    _PICI = value;
                    RaisePropertyChanged("PICI", true);
                }

            }
        }

        [Query]
        [DisplayName("烧结矿")]
        public System.Nullable<double> JS
        {
            get
            {
                return this._JS;
            }
            set
            {
                if (!_JS.Equals(value))
                {
                    _JS = value;
                    RaisePropertyChanged("JS", true);
                }

            }
        }

        [Query]
        [DisplayName("球团矿")]
        public System.Nullable<double> SQ
        {
            get
            {
                return this._SQ;
            }
            set
            {
                if (!_SQ.Equals(value))
                {
                    _SQ = value;
                    RaisePropertyChanged("SQ", true);
                }

            }
        }

        [Query]
        [DisplayName("其它")]
        public System.Nullable<double> QT
        {
            get
            {
                return this._QT;
            }
            set
            {
                if (!_QT.Equals(value))
                {
                    _QT = value;
                    RaisePropertyChanged("QT", true);
                }

            }
        }

        [Query]
        [DisplayName("焦炭")]
        public System.Nullable<double> JT
        {
            get
            {
                return this._JT;
            }
            set
            {
                if (!_JT.Equals(value))
                {
                    _JT = value;
                    RaisePropertyChanged("JT", true);
                }

            }
        }

        [Query]
        [DisplayName("焦丁")]
        public System.Nullable<double> JD
        {
            get
            {
                return this._JD;
            }
            set
            {
                if (!_JD.Equals(value))
                {
                    _JD = value;
                    RaisePropertyChanged("JD", true);
                }

            }
        }

        [Query]
        [DisplayName("矿批")]
        public System.Nullable<double> KP
        {
            get
            {
                return this._KP;
            }
            set
            {
                if (!_KP.Equals(value))
                {
                    _KP = value;
                    RaisePropertyChanged("KP", true);
                }

            }
        }

        [Query]
        [DisplayName("焦批")]
        public System.Nullable<double> JP
        {
            get
            {
                return this._JP;
            }
            set
            {
                if (!_JP.Equals(value))
                {
                    _JP = value;
                    RaisePropertyChanged("JP", true);
                }

            }
        }

        [Query]
        [DisplayName("焦炭负荷")]
        public System.Nullable<double> JTFH
        {
            get
            {
                return this._JTFH;
            }
            set
            {
                if (!_JTFH.Equals(value))
                {
                    _JTFH = value;
                    RaisePropertyChanged("JTFH", true);
                }

            }
        }

        [Query]
        [DisplayName("喷煤量")]
        public System.Nullable<double> PML
        {
            get
            {
                return this._PML;
            }
            set
            {
                if (!_PML.Equals(value))
                {
                    _PML = value;
                    RaisePropertyChanged("PML", true);
                }

            }
        }
        
        [Query]
        [DisplayName("综合负荷")]
        public System.Nullable<double> ZHFH
        {
            get
            {
                return this._ZHFH;
            }
            set
            {
                if (!_ZHFH.Equals(value))
                {
                    _ZHFH = value;
                    RaisePropertyChanged("ZHFH", true);
                }

            }
        }


        public System.Nullable<double> ZL
        {
            get
            {
                return this._ZL;
            }
            set
            {
                if (!_ZL.Equals(value))
                {
                    _ZL = value;
                    RaisePropertyChanged("ZL", true);
                }

            }
        }
        [Query]
        [DisplayName("碱度")]
        public System.Nullable<double> R2
        {
            get
            {
                return this._R2;
            }
            set
            {
                if (!_R2.Equals(value))
                {
                    _R2 = value;
                    RaisePropertyChanged("R2", true);
                }

            }
        }
        [Query]
        [DisplayName("理论碱度")]
        public System.Nullable<double> LLJD
        {
            get
            {
                return this._LLJD;
            }
            set
            {
                if (!_LLJD.Equals(value))
                {
                    _LLJD = value;
                    RaisePropertyChanged("LLJD", true);
                }

            }
        }
        [Query]
        [DisplayName("装料顺序")]
        public string ZLSX
        {
            get
            {
                return this._ZLSX;
            }
            set
            {
                if (!_ZLSX.Equals(value))
                {
                    _ZLSX = value;
                    RaisePropertyChanged("ZLSX", true);
                }

            }
        }

        [Query]
        [DisplayName("料线")]
        public System.Nullable<int> LX
        {
            get
            {
                return this._LX;
            }
            set
            {
                if (!_LX.Equals(value))
                {
                    _LX = value;
                    RaisePropertyChanged("LX", true);
                }

            }
        }

        [Query]
        [DisplayName("变料时间")]
        public System.Nullable<System.DateTime> BLSJ
        {
            get
            {
                return this._BLSJ;
            }
            set
            {
                if (!_BLSJ.Equals(value))
                {
                    _BLSJ = value;
                    RaisePropertyChanged("BLSJ", true);
                }

            }
        }
        [Query]
        [DisplayName("变料原因")]
        public string BLYY
        {
            get
            {
                return this._BLYY;
            }
            set
            {
                if (!_BLYY.Equals(value))
                {
                    _BLYY = value;
                    RaisePropertyChanged("BLYY", true);
                }

            }
        }

        [Query]
        [DisplayName("吃料仓号")]
        public string CLCH
        {
            get
            {
                return this._CLCH;
            }
            set
            {
                if (!_CLCH.Equals(value))
                {
                    _CLCH = value;
                    RaisePropertyChanged("CLCH", true);
                }

            }
        }
        
        [Query]
        [DisplayName("理论铁量")]
        public System.Nullable<double> LLTL
        {
            get
            {
                return this._LLTL;
            }
            set
            {
                if (!_LLTL.Equals(value))
                {
                    _LLTL = value;
                    RaisePropertyChanged("LLTL", true);
                }

            }
        }
        [Query]
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
        [Query]
        [DisplayName("冶炼强度")]
        public System.Nullable<double> YQ
        {
            get
            {
                return this._YQ;
            }
            set
            {
                if (!_YQ.Equals(value))
                {
                    _YQ = value;
                    RaisePropertyChanged("YQ", true);
                }

            }
        }
        [Query]
        [DisplayName("理论班产")]
        public System.Nullable<double> LLBC
        {
            get
            {
                return this._LLBC;
            }
            set
            {
                if (!_LLBC.Equals(value))
                {
                    _LLBC = value;
                    RaisePropertyChanged("LLBC", true);
                }

            }
        }
        [Query]
        [DisplayName("变料角度")]
        public string BLJD
        {
            get
            {
                return this._BLJD;
            }
            set
            {
                if (!_BLJD.Equals(value))
                {
                    _BLJD = value;
                    RaisePropertyChanged("BLJD", true);
                }

            }
        }
        
        public System.Nullable<double> LLJB
        {
            get
            {
                return this._LLJB;
            }
            set
            {
                if (!_LLJB.Equals(value))
                {
                    _LLJB = value;
                    RaisePropertyChanged("LLJB", true);
                }

            }
        }
        
        public System.Nullable<double> SJKPB
        {
            get
            {
                return this._SJKPB;
            }
            set
            {
                if (!_SJKPB.Equals(value))
                {
                    _SJKPB = value;
                    RaisePropertyChanged("SJKPB", true);
                }

            }
        }
        
        public System.Nullable<double> SLQPB
        {
            get
            {
                return this._SLQPB;
            }
            set
            {
                if (!_SLQPB.Equals(value))
                {
                    _SLQPB = value;
                    RaisePropertyChanged("SLQPB", true);
                }

            }
        }
        
        public System.Nullable<double> KKPB
        {
            get
            {
                return this._KKPB;
            }
            set
            {
                if (!_KKPB.Equals(value))
                {
                    _KKPB = value;
                    RaisePropertyChanged("KKPB", true);
                }

            }
        }
        
        public System.Nullable<double> KK
        {
            get
            {
                return this._KK;
            }
            set
            {
                if (!_KK.Equals(value))
                {
                    _KK = value;
                    RaisePropertyChanged("KK", true);
                }

            }
        }
        
        public System.Nullable<double> QRPB
        {
            get
            {
                return this._QRPB;
            }
            set
            {
                if (!_QRPB.Equals(value))
                {
                    _QRPB = value;
                    RaisePropertyChanged("QRPB", true);
                }

            }
        }
        [Query]
        [DisplayName("入炉品位")]
        public System.Nullable<double> RLPW
        {
            get
            {
                return this._RLPW;
            }
            set
            {
                if (!_RLPW.Equals(value))
                {
                    _RLPW = value;
                    RaisePropertyChanged("RLPW", true);
                }

            }
        }
        public System.Nullable<double> BY1
        {
            get
            {
                return this._BY1;
            }
            set
            {
                if (!_BY1.Equals(value))
                {
                    _BY1 = value;
                    RaisePropertyChanged("BY1", true);
                }

            }
        }
        public System.Nullable<double> BY2
        {
            get
            {
                return this._BY2;
            }
            set
            {
                if (!_BY2.Equals(value))
                {
                    _BY2 = value;
                    RaisePropertyChanged("BY2", true);
                }

            }
        }
        
        public override void DbInsert(OracleTransaction trans)
        {
            //InsertSql,InsertPara,AutoSql
            if (DataState == DataStateType.Add)
            {
                OracleConnection con = trans.Connection;
                OracleCommand insertCmd = new OracleCommand();
                insertCmd.Connection = con;
                insertCmd.Transaction = trans;
                insertCmd.CommandText = "INSERT INTO RBZLBD(GAOLU,SJ,BC,PICI,JS,SQ,QT,JT,JD,KP,JP,JTFH,PML,ZHFH,ZL,R2,LLJD,ZLSX,LX,BLSJ,BLYY,CLCH,LLTL,BZ,YQ,LLBC,BLJD,LLJB,SJKPB,SLQPB,KKPB,KK,QRPB,RLPW,BY1,BY2) VALUES(:GAOLU,:SJ,:BC,:PICI,:JS,:SQ,:QT,:JT,:JD,:KP,:JP,:JTFH,:PML,:ZHFH,:ZL,:R2,:LLJD,:ZLSX,:LX,:BLSJ,:BLYY,:CLCH,:LLTL,:BZ,:YQ,:LLBC,:BLJD,:LLJB,:SJKPB,:SLQPB,:KKPB,:KK,:QRPB,:RLPW,:BY1,:BY2)";
                insertCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = (object)this.GAOLU ?? DBNull.Value;
                insertCmd.Parameters.Add(":SJ", OracleType.DateTime, 7).Value = (object)this.SJ ?? DBNull.Value;
                insertCmd.Parameters.Add(":BC", OracleType.VarChar, 10).Value = this.BC;
                insertCmd.Parameters.Add(":PICI", OracleType.Double, 22).Value = (object)this.PICI ?? DBNull.Value;
                insertCmd.Parameters.Add(":JS", OracleType.Double, 22).Value = (object)this.JS ?? DBNull.Value;
                insertCmd.Parameters.Add(":SQ", OracleType.Double, 22).Value = (object)this.SQ ?? DBNull.Value;
                insertCmd.Parameters.Add(":QT", OracleType.Double, 22).Value = (object)this.QT ?? DBNull.Value;
                insertCmd.Parameters.Add(":JT", OracleType.Double, 22).Value = (object)this.JT ?? DBNull.Value;
                insertCmd.Parameters.Add(":JD", OracleType.Double, 22).Value = (object)this.JD ?? DBNull.Value;
                insertCmd.Parameters.Add(":KP", OracleType.Double, 22).Value = (object)this.KP ?? DBNull.Value;
                insertCmd.Parameters.Add(":JP", OracleType.Double, 22).Value = (object)this.JP ?? DBNull.Value;
                insertCmd.Parameters.Add(":JTFH", OracleType.Double, 22).Value = (object)this.JTFH ?? DBNull.Value;
                insertCmd.Parameters.Add(":PML", OracleType.Double, 22).Value = (object)this.PML ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZHFH", OracleType.Double, 22).Value = (object)this.ZHFH ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZL", OracleType.Double, 22).Value = (object)this.ZL ?? DBNull.Value;
                insertCmd.Parameters.Add(":R2", OracleType.Double, 22).Value = (object)this.R2 ?? DBNull.Value;
                insertCmd.Parameters.Add(":LLJD", OracleType.Double, 22).Value = (object)this.LLJD ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZLSX", OracleType.VarChar, 200).Value = this.ZLSX;
                insertCmd.Parameters.Add(":LX", OracleType.Int32, 22).Value = (object)this.LX ?? DBNull.Value;
                insertCmd.Parameters.Add(":BLSJ", OracleType.DateTime, 7).Value = (object)this.BLSJ ?? DBNull.Value;
                insertCmd.Parameters.Add(":BLYY", OracleType.VarChar, 200).Value = this.BLYY;
                insertCmd.Parameters.Add(":CLCH", OracleType.VarChar, 200).Value = this.CLCH;
                insertCmd.Parameters.Add(":LLTL", OracleType.Double, 22).Value = (object)this.LLTL ?? DBNull.Value;
                insertCmd.Parameters.Add(":BZ", OracleType.VarChar, 200).Value = this.BZ;
                insertCmd.Parameters.Add(":YQ", OracleType.Double, 22).Value = (object)this.YQ ?? DBNull.Value;
                insertCmd.Parameters.Add(":LLBC", OracleType.Double, 22).Value = (object)this.LLBC ?? DBNull.Value;
                insertCmd.Parameters.Add(":BLJD", OracleType.VarChar, 300).Value = this.BLJD;
                insertCmd.Parameters.Add(":LLJB", OracleType.Double, 22).Value = (object)this.LLJB ?? DBNull.Value;
                insertCmd.Parameters.Add(":SJKPB", OracleType.Double, 22).Value = (object)this.SJKPB ?? DBNull.Value;
                insertCmd.Parameters.Add(":SLQPB", OracleType.Double, 22).Value = (object)this.SLQPB ?? DBNull.Value;
                insertCmd.Parameters.Add(":KKPB", OracleType.Double, 22).Value = (object)this.KKPB ?? DBNull.Value;
                insertCmd.Parameters.Add(":KK", OracleType.Double, 22).Value = (object)this.KK ?? DBNull.Value;
                insertCmd.Parameters.Add(":QRPB", OracleType.Double, 22).Value = (object)this.QRPB ?? DBNull.Value;
                insertCmd.Parameters.Add(":RLPW", OracleType.Double, 22).Value = (object)this.RLPW ?? DBNull.Value;
                insertCmd.Parameters.Add(":BY1", OracleType.Double, 22).Value = (object)this.BY1 ?? DBNull.Value;
                insertCmd.Parameters.Add(":BY2", OracleType.Double, 22).Value = (object)this.BY2 ?? DBNull.Value;

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
                deleteCmd.CommandText = "DELETE FROM RBZLBD WHERE ROWID='" + RowId.Value + "'";

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
                updateCmd.CommandText = "UPDATE RBZLBD SET GAOLU=:GAOLU,SJ=:SJ,BC=:BC,PICI=:PICI,JS=:JS,SQ=:SQ,QT=:QT,JT=:JT,JD=:JD,KP=:KP,JP=:JP,JTFH=:JTFH,PML=:PML,ZHFH=:ZHFH,ZL=:ZL,R2=:R2,LLJD=:LLJD,ZLSX=:ZLSX,LX=:LX,BLSJ=:BLSJ,BLYY=:BLYY,CLCH=:CLCH,LLTL=:LLTL,BZ=:BZ,YQ=:YQ,LLBC=:LLBC,BLJD=:BLJD,LLJB=:LLJB,SJKPB=:SJKPB,SLQPB=:SLQPB,KKPB=:KKPB,KK=:KK,QRPB=:QRPB,RLPW=:RLPW ,BY1=:BY1 ,BY2=:BY2  WHERE ROWID='" + RowId.Value + "'";
                updateCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = (object)this.GAOLU ?? DBNull.Value;
                updateCmd.Parameters.Add(":SJ", OracleType.DateTime, 7).Value = (object)this.SJ ?? DBNull.Value;
                updateCmd.Parameters.Add(":BC", OracleType.VarChar, 10).Value = this.BC;
                updateCmd.Parameters.Add(":PICI", OracleType.Double, 22).Value = (object)this.PICI ?? DBNull.Value;
                updateCmd.Parameters.Add(":JS", OracleType.Double, 22).Value = (object)this.JS ?? DBNull.Value;
                updateCmd.Parameters.Add(":SQ", OracleType.Double, 22).Value = (object)this.SQ ?? DBNull.Value;
                updateCmd.Parameters.Add(":QT", OracleType.Double, 22).Value = (object)this.QT ?? DBNull.Value;
                updateCmd.Parameters.Add(":JT", OracleType.Double, 22).Value = (object)this.JT ?? DBNull.Value;
                updateCmd.Parameters.Add(":JD", OracleType.Double, 22).Value = (object)this.JD ?? DBNull.Value;
                updateCmd.Parameters.Add(":KP", OracleType.Double, 22).Value = (object)this.KP ?? DBNull.Value;
                updateCmd.Parameters.Add(":JP", OracleType.Double, 22).Value = (object)this.JP ?? DBNull.Value;
                updateCmd.Parameters.Add(":JTFH", OracleType.Double, 22).Value = (object)this.JTFH ?? DBNull.Value;
                updateCmd.Parameters.Add(":PML", OracleType.Double, 22).Value = (object)this.PML ?? DBNull.Value;
                updateCmd.Parameters.Add(":ZHFH", OracleType.Double, 22).Value = (object)this.ZHFH ?? DBNull.Value;
                updateCmd.Parameters.Add(":ZL", OracleType.Double, 22).Value = (object)this.ZL ?? DBNull.Value;
                updateCmd.Parameters.Add(":R2", OracleType.Double, 22).Value = (object)this.R2 ?? DBNull.Value;
                updateCmd.Parameters.Add(":LLJD", OracleType.Double, 22).Value = (object)this.LLJD ?? DBNull.Value;
                updateCmd.Parameters.Add(":ZLSX", OracleType.VarChar, 200).Value = this.ZLSX;
                updateCmd.Parameters.Add(":LX", OracleType.Int32, 22).Value = (object)this.LX ?? DBNull.Value;
                updateCmd.Parameters.Add(":BLSJ", OracleType.DateTime, 7).Value = (object)this.BLSJ ?? DBNull.Value;
                updateCmd.Parameters.Add(":BLYY", OracleType.VarChar, 200).Value = this.BLYY;
                updateCmd.Parameters.Add(":CLCH", OracleType.VarChar, 200).Value = this.CLCH;
                updateCmd.Parameters.Add(":LLTL", OracleType.Double, 22).Value = (object)this.LLTL ?? DBNull.Value;
                updateCmd.Parameters.Add(":BZ", OracleType.VarChar, 200).Value = this.BZ;
                updateCmd.Parameters.Add(":YQ", OracleType.Double, 22).Value = (object)this.YQ ?? DBNull.Value;
                updateCmd.Parameters.Add(":LLBC", OracleType.Double, 22).Value = (object)this.LLBC ?? DBNull.Value;
                updateCmd.Parameters.Add(":BLJD", OracleType.VarChar, 300).Value = this.BLJD;
                updateCmd.Parameters.Add(":LLJB", OracleType.Double, 22).Value = (object)this.LLJB ?? DBNull.Value;
                updateCmd.Parameters.Add(":SJKPB", OracleType.Double, 22).Value = (object)this.SJKPB ?? DBNull.Value;
                updateCmd.Parameters.Add(":SLQPB", OracleType.Double, 22).Value = (object)this.SLQPB ?? DBNull.Value;
                updateCmd.Parameters.Add(":KKPB", OracleType.Double, 22).Value = (object)this.KKPB ?? DBNull.Value;
                updateCmd.Parameters.Add(":KK", OracleType.Double, 22).Value = (object)this.KK ?? DBNull.Value;
                updateCmd.Parameters.Add(":QRPB", OracleType.Double, 22).Value = (object)this.QRPB ?? DBNull.Value;
                updateCmd.Parameters.Add(":RLPW", OracleType.Double, 22).Value = (object)this.RLPW ?? DBNull.Value;
                  updateCmd.Parameters.Add(":BY1", OracleType.Double, 22).Value = (object)this.BY1 ?? DBNull.Value;
                  updateCmd.Parameters.Add(":BY2", OracleType.Double, 22).Value = (object)this.BY2 ?? DBNull.Value;

                updateCmd.ExecuteNonQuery();
            }


        }
    }
    public partial class RbzlbdTable : DbEntityTable<Rbzlbd>, IFilter
    {

        private DateTime? loadRq = null;
        private int? loadGaolu = null;
        //Rbzlbd,ltznConnectionString
        public void LoadByRQGAOLU(DateTime rq,int gaolu)
        {
            //ConStr, EntiyType, SelectSql, SelectPara, SelectParaVal
            loadRq = rq;
            loadGaolu = gaolu;
            beginAdjustData = true;
            this.Clear();
            hideList.Clear();
            OracleConnection con = new OracleConnection(LTZN.Properties.Settings.Default.ltznConnectionString);
            con.Open();
            OracleCommand selectCmd = new OracleCommand();
            selectCmd.Connection = con;
            selectCmd.CommandText = "SELECT GAOLU,SJ,BC,PICI,JS,SQ,QT,JT,JD,KP,JP,JTFH,PML,ZHFH,ZL,R2,LLJD,ZLSX,LX,BLSJ,BLYY,CLCH,LLTL,BZ,YQ,LLBC,BLJD,LLJB,SJKPB,SLQPB,KKPB,KK,QRPB,RLPW,BY1,BY2,ROWID FROM RBZLBD WHERE TRUNC(SJ)=:RQ and gaolu=:GAOLU order by blsj";
            selectCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = rq;
            selectCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = gaolu;

            OracleDataReader dr = selectCmd.ExecuteReader();
            while (dr.Read())
            {
                Rbzlbd item = new Rbzlbd();
                item.BeginInit();
                if (dr.IsDBNull(0)) item.GAOLU = null; else item.GAOLU = dr.GetInt32(0);
                if (dr.IsDBNull(1)) item.SJ = null; else item.SJ = dr.GetDateTime(1);
                if (dr.IsDBNull(2)) item.BC = ""; else item.BC = dr.GetString(2);
                if (dr.IsDBNull(3)) item.PICI = null; else item.PICI = dr.GetDouble(3);
                if (dr.IsDBNull(4)) item.JS = null; else item.JS = dr.GetDouble(4);
                if (dr.IsDBNull(5)) item.SQ = null; else item.SQ = dr.GetDouble(5);
                if (dr.IsDBNull(6)) item.QT = null; else item.QT = dr.GetDouble(6);
                if (dr.IsDBNull(7)) item.JT = null; else item.JT = dr.GetDouble(7);
                if (dr.IsDBNull(8)) item.JD = null; else item.JD = dr.GetDouble(8);
                if (dr.IsDBNull(9)) item.KP = null; else item.KP = dr.GetDouble(9);
                if (dr.IsDBNull(10)) item.JP = null; else item.JP = dr.GetDouble(10);
                if (dr.IsDBNull(11)) item.JTFH = null; else item.JTFH = dr.GetDouble(11);
                if (dr.IsDBNull(12)) item.PML = null; else item.PML = dr.GetDouble(12);
                if (dr.IsDBNull(13)) item.ZHFH = null; else item.ZHFH = dr.GetDouble(13);
                if (dr.IsDBNull(14)) item.ZL = null; else item.ZL = dr.GetDouble(14);
                if (dr.IsDBNull(15)) item.R2 = null; else item.R2 = dr.GetDouble(15);
                if (dr.IsDBNull(16)) item.LLJD = null; else item.LLJD = dr.GetDouble(16);
                if (dr.IsDBNull(17)) item.ZLSX = ""; else item.ZLSX = dr.GetString(17);
                if (dr.IsDBNull(18)) item.LX = null; else item.LX = dr.GetInt32(18);
                if (dr.IsDBNull(19)) item.BLSJ = null; else item.BLSJ = dr.GetDateTime(19);
                if (dr.IsDBNull(20)) item.BLYY = ""; else item.BLYY = dr.GetString(20);
                if (dr.IsDBNull(21)) item.CLCH = ""; else item.CLCH = dr.GetString(21);
                if (dr.IsDBNull(22)) item.LLTL = null; else item.LLTL = dr.GetDouble(22);
                if (dr.IsDBNull(23)) item.BZ = ""; else item.BZ = dr.GetString(23);
                if (dr.IsDBNull(24)) item.YQ = null; else item.YQ = dr.GetDouble(24);
                if (dr.IsDBNull(25)) item.LLBC = null; else item.LLBC = dr.GetDouble(25);
                if (dr.IsDBNull(26)) item.BLJD = ""; else item.BLJD = dr.GetString(26);
                if (dr.IsDBNull(27)) item.LLJB = null; else item.LLJB = dr.GetDouble(27);
                if (dr.IsDBNull(28)) item.SJKPB = null; else item.SJKPB = dr.GetDouble(28);
                if (dr.IsDBNull(29)) item.SLQPB = null; else item.SLQPB = dr.GetDouble(29);
                if (dr.IsDBNull(30)) item.KKPB = null; else item.KKPB = dr.GetDouble(30);
                if (dr.IsDBNull(31)) item.KK = null; else item.KK = dr.GetDouble(31);
                if (dr.IsDBNull(32)) item.QRPB = null; else item.QRPB = dr.GetDouble(32);
                if (dr.IsDBNull(33)) item.RLPW = null; else item.RLPW = dr.GetDouble(33);
                if (dr.IsDBNull(34)) item.BY1 = null; else item.BY1 = dr.GetDouble(34);
                if (dr.IsDBNull(35)) item.BY2 = null; else item.BY2 = dr.GetDouble(35);
                item.RowId = dr.GetString(36);

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

        public void Load(string whereSql)
        {
            //ConStr, EntiyType, SelectSql, SelectPara, SelectParaVal
            beginAdjustData = true;
            this.Clear();
            hideList.Clear();
            OracleConnection con = new OracleConnection(LTZN.Properties.Settings.Default.ltznConnectionString);
            con.Open();
            OracleCommand selectCmd = new OracleCommand();
            selectCmd.Connection = con;
            selectCmd.CommandText = "SELECT GAOLU,SJ,BC,PICI,JS,SQ,QT,JT,JD,KP,JP,JTFH,PML,ZHFH,ZL,R2,LLJD,ZLSX,LX,BLSJ,BLYY,CLCH,LLTL,BZ,YQ,LLBC,BLJD,LLJB,SJKPB,SLQPB,KKPB,KK,QRPB,RLPW,BY1,BY2,ROWID FROM RBZLBD WHERE " + whereSql + " order by GAOLU,blsj";

            OracleDataReader dr = selectCmd.ExecuteReader();
            while (dr.Read())
            {
                Rbzlbd item = new Rbzlbd();
                item.BeginInit();
                if (dr.IsDBNull(0)) item.GAOLU = null; else item.GAOLU = dr.GetInt32(0);
                if (dr.IsDBNull(1)) item.SJ = null; else item.SJ = dr.GetDateTime(1);
                if (dr.IsDBNull(2)) item.BC = ""; else item.BC = dr.GetString(2);
                if (dr.IsDBNull(3)) item.PICI = null; else item.PICI = dr.GetDouble(3);
                if (dr.IsDBNull(4)) item.JS = null; else item.JS = dr.GetDouble(4);
                if (dr.IsDBNull(5)) item.SQ = null; else item.SQ = dr.GetDouble(5);
                if (dr.IsDBNull(6)) item.QT = null; else item.QT = dr.GetDouble(6);
                if (dr.IsDBNull(7)) item.JT = null; else item.JT = dr.GetDouble(7);
                if (dr.IsDBNull(8)) item.JD = null; else item.JD = dr.GetDouble(8);
                if (dr.IsDBNull(9)) item.KP = null; else item.KP = dr.GetDouble(9);
                if (dr.IsDBNull(10)) item.JP = null; else item.JP = dr.GetDouble(10);
                if (dr.IsDBNull(11)) item.JTFH = null; else item.JTFH = dr.GetDouble(11);
                if (dr.IsDBNull(12)) item.PML = null; else item.PML = dr.GetDouble(12);
                if (dr.IsDBNull(13)) item.ZHFH = null; else item.ZHFH = dr.GetDouble(13);
                if (dr.IsDBNull(14)) item.ZL = null; else item.ZL = dr.GetDouble(14);
                if (dr.IsDBNull(15)) item.R2 = null; else item.R2 = dr.GetDouble(15);
                if (dr.IsDBNull(16)) item.LLJD = null; else item.LLJD = dr.GetDouble(16);
                if (dr.IsDBNull(17)) item.ZLSX = ""; else item.ZLSX = dr.GetString(17);
                if (dr.IsDBNull(18)) item.LX = null; else item.LX = dr.GetInt32(18);
                if (dr.IsDBNull(19)) item.BLSJ = null; else item.BLSJ = dr.GetDateTime(19);
                if (dr.IsDBNull(20)) item.BLYY = ""; else item.BLYY = dr.GetString(20);
                if (dr.IsDBNull(21)) item.CLCH = ""; else item.CLCH = dr.GetString(21);
                if (dr.IsDBNull(22)) item.LLTL = null; else item.LLTL = dr.GetDouble(22);
                if (dr.IsDBNull(23)) item.BZ = ""; else item.BZ = dr.GetString(23);
                if (dr.IsDBNull(24)) item.YQ = null; else item.YQ = dr.GetDouble(24);
                if (dr.IsDBNull(25)) item.LLBC = null; else item.LLBC = dr.GetDouble(25);
                if (dr.IsDBNull(26)) item.BLJD = ""; else item.BLJD = dr.GetString(26);
                if (dr.IsDBNull(27)) item.LLJB = null; else item.LLJB = dr.GetDouble(27);
                if (dr.IsDBNull(28)) item.SJKPB = null; else item.SJKPB = dr.GetDouble(28);
                if (dr.IsDBNull(29)) item.SLQPB = null; else item.SLQPB = dr.GetDouble(29);
                if (dr.IsDBNull(30)) item.KKPB = null; else item.KKPB = dr.GetDouble(30);
                if (dr.IsDBNull(31)) item.KK = null; else item.KK = dr.GetDouble(31);
                if (dr.IsDBNull(32)) item.QRPB = null; else item.QRPB = dr.GetDouble(32);
                if (dr.IsDBNull(33)) item.RLPW = null; else item.RLPW = dr.GetDouble(33);
                if (dr.IsDBNull(34)) item.BY1 = null; else item.BY1 = dr.GetDouble(34);
                if (dr.IsDBNull(35)) item.BY2 = null; else item.BY2 = dr.GetDouble(35);
                item.RowId = dr.GetString(36);

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

        public override void Save()
        {
            if (loadRq != null && loadGaolu != null)
            {
                int i = 0;
                foreach (var item in this)
                {

                    if (item.SJ == null)
                        item.SJ = DateTime.Now.AddMinutes(i++);

                    item.SJ = new DateTime(loadRq.Value.Year, loadRq.Value.Month, loadRq.Value.Day, item.SJ.Value.Hour, item.SJ.Value.Minute, 0);

                    item.GAOLU = loadGaolu.Value;
                }
            }
            base.Save();
        }
    }
}
