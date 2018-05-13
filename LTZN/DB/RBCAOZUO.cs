using System;
using System.ComponentModel;
using System.Data.OracleClient;

namespace ZHCDB
{
    public partial class Rbcaozuo : DbEntity
    {
        private System.Nullable<int> _FJ;
        private double _LFLL;
        private System.Nullable<double> _FYLL;
        private System.Nullable<double> _LFYL;
        private System.Nullable<double> _RFYL;
        private System.Nullable<double> _RFWD;
        private System.Nullable<double> _PML;
        private System.Nullable<double> _TQX;
        private System.Nullable<double> _LDWD;
        private System.Nullable<double> _RSWD;
        private System.Nullable<double> _SJFS;
        private System.Nullable<double> _GFDN;
        private System.Nullable<double> _LDYL;
        private System.Nullable<double> _LX;
        private string _BZ = "";
        private int _GAOLU;
        private System.Nullable<System.DateTime> _SJ = null;
        private int _HOUR;
        private System.Nullable<double> _ZHFH;
        private string _BANBIE = "";
        private System.Nullable<double> _MIANJI;
        private System.Nullable<double> _FYL;
        //ConStr 

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


        public System.Nullable<int> FJ
        {
            get
            {
                return this._FJ;
            }
            set
            {
                if (!_FJ.Equals(value))
                {
                    _FJ = value;
                    RaisePropertyChanged("FJ", true);
                }

            }
        }

        [Query]
        [DisplayName("冷风流量")]
        public double LFLL
        {
            get
            {
                return this._LFLL;
            }
            set
            {
                if (!_LFLL.Equals(value))
                {
                    _LFLL = value;
                    RaisePropertyChanged("LFLL", true);
                }

            }
        }

        [Query]
        [DisplayName("富氧流量")]
        public System.Nullable<double> FYLL
        {
            get
            {
                return this._FYLL;
            }
            set
            {
                if (!_FYLL.Equals(value))
                {
                    _FYLL = value;
                    RaisePropertyChanged("FYLL", true);
                }

            }
        }

        [Query]
        [DisplayName("冷风压力")]
        public System.Nullable<double> LFYL
        {
            get
            {
                return this._LFYL;
            }
            set
            {
                if (!_LFYL.Equals(value))
                {
                    _LFYL = value;
                    RaisePropertyChanged("LFYL", true);
                }

            }
        }

        [Query]
        [DisplayName("热风压力")]
        public System.Nullable<double> RFYL
        {
            get
            {
                return this._RFYL;
            }
            set
            {
                if (!_RFYL.Equals(value))
                {
                    _RFYL = value;
                    RaisePropertyChanged("RFYL", true);
                }

            }
        }

        [Query]
        [DisplayName("热风温度")] 
        public System.Nullable<double> RFWD
        {
            get
            {
                return this._RFWD;
            }
            set
            {
                if (!_RFWD.Equals(value))
                {
                    _RFWD = value;
                    RaisePropertyChanged("RFWD", true);
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
        [DisplayName("透气性")]
        public System.Nullable<double> TQX
        {
            get
            {
                return this._TQX;
            }
            set
            {
                if (!_TQX.Equals(value))
                {
                    _TQX = value;
                    RaisePropertyChanged("TQX", true);
                }

            }
        }

        [Query]
        [DisplayName("炉顶温度")]
        public System.Nullable<double> LDWD
        {
            get
            {
                return this._LDWD;
            }
            set
            {
                if (!_LDWD.Equals(value))
                {
                    _LDWD = value;
                    RaisePropertyChanged("LDWD", true);
                }

            }
        }

        [Query]
        [DisplayName("燃烧温度")]
        public System.Nullable<double> RSWD
        {
            get
            {
                return this._RSWD;
            }
            set
            {
                if (!_RSWD.Equals(value))
                {
                    _RSWD = value;
                    RaisePropertyChanged("RSWD", true);
                }

            }
        }

        [Query]
        [DisplayName("实际风速")]
        public System.Nullable<double> SJFS
        {
            get
            {
                return this._SJFS;
            }
            set
            {
                if (!_SJFS.Equals(value))
                {
                    _SJFS = value;
                    RaisePropertyChanged("SJFS", true);
                }

            }
        }

        [Query]
        [DisplayName("鼓风动能")]
        public System.Nullable<double> GFDN
        {
            get
            {
                return this._GFDN;
            }
            set
            {
                if (!_GFDN.Equals(value))
                {
                    _GFDN = value;
                    RaisePropertyChanged("GFDN", true);
                }

            }
        }

        [Query]
        [DisplayName("炉顶压力")]
        public System.Nullable<double> LDYL
        {
            get
            {
                return this._LDYL;
            }
            set
            {
                if (!_LDYL.Equals(value))
                {
                    _LDYL = value;
                    RaisePropertyChanged("LDYL", true);
                }

            }
        }

        [Query]
        [DisplayName("料批")]
        public System.Nullable<double> LX
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
        
        
        public int HOUR
        {
            get
            {
                return this._HOUR;
            }
            set
            {
                if (!_HOUR.Equals(value))
                {
                    _HOUR = value;
                    RaisePropertyChanged("HOUR", true);
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
        public string BANBIE
        {
            get
            {
                return this._BANBIE;
            }
            set
            {
                if (!_BANBIE.Equals(value))
                {
                    _BANBIE = value;
                    RaisePropertyChanged("BANBIE", true);
                }

            }
        }
        public System.Nullable<double> MIANJI
        {
            get
            {
                return this._MIANJI;
            }
            set
            {
                if (!_MIANJI.Equals(value))
                {
                    _MIANJI = value;
                    RaisePropertyChanged("MIANJI", true);
                }

            }
        }
        public System.Nullable<double> FYL
        {
            get
            {
                return (this.FYLL * 0.785*100) / (60 * this.LFLL);
            }
            set
            {
                if (!_FYL.Equals(value))
                {
                    _FYL = value;
                    RaisePropertyChanged("FYL", true);
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
                insertCmd.CommandText = "INSERT INTO RBCAOZUO(FJ,LFLL,FYLL,LFYL,RFYL,RFWD,PML,TQX,LDWD,RSWD,SJFS,GFDN,LDYL,LX,BZ,GAOLU,SJ,HOUR,ZHFH,BANBIE,MIANJI,FYL) VALUES(:FJ,:LFLL,:FYLL,:LFYL,:RFYL,:RFWD,:PML,:TQX,:LDWD,:RSWD,:SJFS,:GFDN,:LDYL,:LX,:BZ,:GAOLU,:SJ,:HOUR,:ZHFH,:BANBIE,:MIANJI,:FYL)";
                insertCmd.Parameters.Add(":FJ", OracleType.Int32, 22).Value = (object)this.FJ ?? DBNull.Value;
                insertCmd.Parameters.Add(":LFLL", OracleType.Double, 22).Value = this.LFLL;
                insertCmd.Parameters.Add(":FYLL", OracleType.Double, 22).Value = (object)this.FYLL ?? DBNull.Value;
                insertCmd.Parameters.Add(":LFYL", OracleType.Double, 22).Value = (object)this.LFYL ?? DBNull.Value;
                insertCmd.Parameters.Add(":RFYL", OracleType.Double, 22).Value = (object)this.RFYL ?? DBNull.Value;
                insertCmd.Parameters.Add(":RFWD", OracleType.Double, 22).Value = (object)this.RFWD ?? DBNull.Value;
                insertCmd.Parameters.Add(":PML", OracleType.Double, 22).Value = (object)this.PML ?? DBNull.Value;
                insertCmd.Parameters.Add(":TQX", OracleType.Double, 22).Value = (object)this.TQX ?? DBNull.Value;
                insertCmd.Parameters.Add(":LDWD", OracleType.Double, 22).Value = (object)this.LDWD ?? DBNull.Value;
                insertCmd.Parameters.Add(":RSWD", OracleType.Double, 22).Value = (object)this.RSWD ?? DBNull.Value;
                insertCmd.Parameters.Add(":SJFS", OracleType.Double, 22).Value = (object)this.SJFS ?? DBNull.Value;
                insertCmd.Parameters.Add(":GFDN", OracleType.Double, 22).Value = (object)this.GFDN ?? DBNull.Value;
                insertCmd.Parameters.Add(":LDYL", OracleType.Double, 22).Value = (object)this.LDYL ?? DBNull.Value;
                insertCmd.Parameters.Add(":LX", OracleType.Double, 22).Value = (object)this.LX ?? DBNull.Value;
                insertCmd.Parameters.Add(":BZ", OracleType.VarChar, 400).Value = this.BZ;
                insertCmd.Parameters.Add(":GAOLU", OracleType.Int32, 22).Value = this.GAOLU;
                insertCmd.Parameters.Add(":SJ", OracleType.DateTime, 7).Value = (object)this.SJ ?? DBNull.Value;
                insertCmd.Parameters.Add(":HOUR", OracleType.Int32, 22).Value = this.HOUR;
                insertCmd.Parameters.Add(":ZHFH", OracleType.Double, 22).Value = (object)this.ZHFH ?? DBNull.Value;
                insertCmd.Parameters.Add(":BANBIE", OracleType.VarChar, 20).Value = this.BANBIE;
                insertCmd.Parameters.Add(":MIANJI", OracleType.Double, 22).Value = (object)this.MIANJI ?? DBNull.Value;
                insertCmd.Parameters.Add(":FYL", OracleType.Double, 22).Value = (object)this.FYL ?? DBNull.Value;
           
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
                deleteCmd.CommandText = "DELETE FROM RBCAOZUO WHERE ROWID='" + RowId.Value + "'";

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
                updateCmd.CommandText = "UPDATE RBCAOZUO SET FJ=:FJ,LFLL=:LFLL,FYLL=:FYLL,LFYL=:LFYL,RFYL=:RFYL,RFWD=:RFWD,PML=:PML,TQX=:TQX,LDWD=:LDWD,RSWD=:RSWD,SJFS=:SJFS,GFDN=:GFDN,LDYL=:LDYL,LX=:LX,BZ=:BZ,GAOLU=:GAOLU,SJ=:SJ,HOUR=:HOUR,ZHFH=:ZHFH,BANBIE=:BANBIE,MIANJI=:MIANJI,FYL=:FYL  WHERE ROWID='" + RowId.Value + "'";
                updateCmd.Parameters.Add(":FJ", OracleType.Int32, 22).Value = (object)this.FJ ?? DBNull.Value;
                updateCmd.Parameters.Add(":LFLL", OracleType.Double, 22).Value = this.LFLL;
                updateCmd.Parameters.Add(":FYLL", OracleType.Double, 22).Value = (object)this.FYLL ?? DBNull.Value;
                updateCmd.Parameters.Add(":LFYL", OracleType.Double, 22).Value = (object)this.LFYL ?? DBNull.Value;
                updateCmd.Parameters.Add(":RFYL", OracleType.Double, 22).Value = (object)this.RFYL ?? DBNull.Value;
                updateCmd.Parameters.Add(":RFWD", OracleType.Double, 22).Value = (object)this.RFWD ?? DBNull.Value;
                updateCmd.Parameters.Add(":PML", OracleType.Double, 22).Value = (object)this.PML ?? DBNull.Value;
                updateCmd.Parameters.Add(":TQX", OracleType.Double, 22).Value = (object)this.TQX ?? DBNull.Value;
                updateCmd.Parameters.Add(":LDWD", OracleType.Double, 22).Value = (object)this.LDWD ?? DBNull.Value;
                updateCmd.Parameters.Add(":RSWD", OracleType.Double, 22).Value = (object)this.RSWD ?? DBNull.Value;
                updateCmd.Parameters.Add(":SJFS", OracleType.Double, 22).Value = (object)this.SJFS ?? DBNull.Value;
                updateCmd.Parameters.Add(":GFDN", OracleType.Double, 22).Value = (object)this.GFDN ?? DBNull.Value;
                updateCmd.Parameters.Add(":LDYL", OracleType.Double, 22).Value = (object)this.LDYL ?? DBNull.Value;
                updateCmd.Parameters.Add(":LX", OracleType.Double, 22).Value = (object)this.LX ?? DBNull.Value;
                updateCmd.Parameters.Add(":BZ", OracleType.VarChar, 400).Value = this.BZ;
                updateCmd.Parameters.Add(":GAOLU", OracleType.Int32, 22).Value = this.GAOLU;
                updateCmd.Parameters.Add(":SJ", OracleType.DateTime, 7).Value = (object)this.SJ ?? DBNull.Value;
                updateCmd.Parameters.Add(":HOUR", OracleType.Int32, 22).Value = this.HOUR;
                updateCmd.Parameters.Add(":ZHFH", OracleType.Double, 22).Value = (object)this.ZHFH ?? DBNull.Value;
                updateCmd.Parameters.Add(":BANBIE", OracleType.VarChar, 20).Value = this.BANBIE;
                updateCmd.Parameters.Add(":MIANJI", OracleType.Double, 22).Value = (object)this.MIANJI ?? DBNull.Value;
                updateCmd.Parameters.Add(":FYL", OracleType.Double, 22).Value = (object)this.FYL ?? DBNull.Value;
                    updateCmd.ExecuteNonQuery();
            }


        }
    }
    public partial class RbcaozuoTable : DbEntityTable<Rbcaozuo>, IFilter
    {
        private DateTime? loadRq = null;
        private int? loadGaolu = null;
        //Rbcaozuo,ltznConnectionString
        public void LoadByRQGAOLU(DateTime rq, int gaolu)
        {
            loadRq = rq;
            loadGaolu = gaolu;
            //ConStr, EntiyType, SelectSql, SelectPara, SelectParaVal
            beginAdjustData = true;
            this.Clear();
            hideList.Clear();
            OracleConnection con = new OracleConnection(LTZN.Properties.Settings.Default.ltznConnectionString);
            con.Open();
            OracleCommand selectCmd = new OracleCommand();
            selectCmd.Connection = con;
            selectCmd.CommandText = "SELECT FJ,LFLL,FYLL,LFYL,RFYL,RFWD,PML,TQX,LDWD,RSWD,SJFS,GFDN,LDYL,LX,BZ,GAOLU,SJ,HOUR,ZHFH,BANBIE,MIANJI,FYL,ROWID FROM RBCAOZUO WHERE TRUNC(SJ)=:RQ and GAOLU=:GAOLU order by hour";
            selectCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = rq;
            selectCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = gaolu;

            OracleDataReader dr = selectCmd.ExecuteReader();
            while (dr.Read())
            {
                Rbcaozuo item = new Rbcaozuo();
                item.BeginInit();
                if (dr.IsDBNull(0)) item.FJ = null; else item.FJ = dr.GetInt32(0);
                if (dr.IsDBNull(1)) item.LFLL = 0; else item.LFLL = dr.GetDouble(1);
                if (dr.IsDBNull(2)) item.FYLL = null; else item.FYLL = dr.GetDouble(2);
                if (dr.IsDBNull(3)) item.LFYL = null; else item.LFYL = dr.GetDouble(3);
                if (dr.IsDBNull(4)) item.RFYL = null; else item.RFYL = dr.GetDouble(4);
                if (dr.IsDBNull(5)) item.RFWD = null; else item.RFWD = dr.GetDouble(5);
                if (dr.IsDBNull(6)) item.PML = null; else item.PML = dr.GetDouble(6);
                if (dr.IsDBNull(7)) item.TQX = null; else item.TQX = dr.GetDouble(7);
                if (dr.IsDBNull(8)) item.LDWD = null; else item.LDWD = dr.GetDouble(8);
                if (dr.IsDBNull(9)) item.RSWD = null; else item.RSWD = dr.GetDouble(9);
                if (dr.IsDBNull(10)) item.SJFS = null; else item.SJFS = dr.GetDouble(10);
                if (dr.IsDBNull(11)) item.GFDN = null; else item.GFDN = dr.GetDouble(11);
                if (dr.IsDBNull(12)) item.LDYL = null; else item.LDYL = dr.GetDouble(12);
                if (dr.IsDBNull(13)) item.LX = null; else item.LX = dr.GetDouble(13);
                if (dr.IsDBNull(14)) item.BZ = ""; else item.BZ = dr.GetString(14);
                if (dr.IsDBNull(15)) item.GAOLU = 0; else item.GAOLU = dr.GetInt32(15);
                if (dr.IsDBNull(16)) item.SJ = null; else item.SJ = dr.GetDateTime(16);
                if (dr.IsDBNull(17)) item.HOUR = 0; else item.HOUR = dr.GetInt32(17);
                if (dr.IsDBNull(18)) item.ZHFH = null; else item.ZHFH = dr.GetDouble(18);
                if (dr.IsDBNull(19)) item.BANBIE = ""; else item.BANBIE = dr.GetString(19);
                if (dr.IsDBNull(20)) item.MIANJI = null; else item.MIANJI = dr.GetDouble(20);
                if (dr.IsDBNull(21)) item.FYL = null; else item.FYL = dr.GetDouble(21);
                
                item.RowId = dr.GetString(22);

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
            beginAdjustData = true;
            this.Clear();
            hideList.Clear();
            OracleConnection con = new OracleConnection(LTZN.Properties.Settings.Default.ltznConnectionString);
            con.Open();
            OracleCommand selectCmd = new OracleCommand();
            selectCmd.Connection = con;
            selectCmd.CommandText = "SELECT FJ,LFLL,FYLL,LFYL,RFYL,RFWD,PML,TQX,LDWD,RSWD,SJFS,GFDN,LDYL,LX,BZ,GAOLU,SJ,HOUR,ZHFH,BANBIE,MIANJI,FYL,ROWID FROM RBCAOZUO WHERE " +whereSql +" Order by GAOLU,SJ";

            OracleDataReader dr = selectCmd.ExecuteReader();
            while (dr.Read())
            {
                Rbcaozuo item = new Rbcaozuo();
                item.BeginInit();
                if (dr.IsDBNull(0)) item.FJ = null; else item.FJ = dr.GetInt32(0);
                if (dr.IsDBNull(1)) item.LFLL = 0; else item.LFLL = dr.GetDouble(1);
                if (dr.IsDBNull(2)) item.FYLL = null; else item.FYLL = dr.GetDouble(2);
                if (dr.IsDBNull(3)) item.LFYL = null; else item.LFYL = dr.GetDouble(3);
                if (dr.IsDBNull(4)) item.RFYL = null; else item.RFYL = dr.GetDouble(4);
                if (dr.IsDBNull(5)) item.RFWD = null; else item.RFWD = dr.GetDouble(5);
                if (dr.IsDBNull(6)) item.PML = null; else item.PML = dr.GetDouble(6);
                if (dr.IsDBNull(7)) item.TQX = null; else item.TQX = dr.GetDouble(7);
                if (dr.IsDBNull(8)) item.LDWD = null; else item.LDWD = dr.GetDouble(8);
                if (dr.IsDBNull(9)) item.RSWD = null; else item.RSWD = dr.GetDouble(9);
                if (dr.IsDBNull(10)) item.SJFS = null; else item.SJFS = dr.GetDouble(10);
                if (dr.IsDBNull(11)) item.GFDN = null; else item.GFDN = dr.GetDouble(11);
                if (dr.IsDBNull(12)) item.LDYL = null; else item.LDYL = dr.GetDouble(12);
                if (dr.IsDBNull(13)) item.LX = null; else item.LX = dr.GetDouble(13);
                if (dr.IsDBNull(14)) item.BZ = ""; else item.BZ = dr.GetString(14);
                if (dr.IsDBNull(15)) item.GAOLU = 0; else item.GAOLU = dr.GetInt32(15);
                if (dr.IsDBNull(16)) item.SJ = null; else item.SJ = dr.GetDateTime(16);
                if (dr.IsDBNull(17)) item.HOUR = 0; else item.HOUR = dr.GetInt32(17);
                if (dr.IsDBNull(18)) item.ZHFH = null; else item.ZHFH = dr.GetDouble(18);
                if (dr.IsDBNull(19)) item.BANBIE = ""; else item.BANBIE = dr.GetString(19);
                if (dr.IsDBNull(20)) item.MIANJI = null; else item.MIANJI = dr.GetDouble(20);
                if (dr.IsDBNull(21)) item.FYL = null; else item.FYL = dr.GetDouble(21);

                item.RowId = dr.GetString(22);

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
