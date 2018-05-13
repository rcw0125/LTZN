using System;
using System.ComponentModel;
using System.Data.OracleClient;

namespace ZHCDB
{
    public partial class Ddmf : DbEntity
    {
        private string _PZ = "无烟煤";
        private System.DateTime _SJ = DateTime.Now;
        private System.Nullable<double> _C;
        private System.Nullable<double> _HUIFEN;
        private System.Nullable<double> _HUIFA;
        private System.Nullable<double> _S;
        private System.Nullable<double> _SF;
        private System.Nullable<double> _XIDU;
        private System.Nullable<double> _FAREZHI;
        //ConStr 



        public string PZ
        {
            get
            {
                return this._PZ;
            }
            set
            {
                if (!_PZ.Equals(value))
                {
                    _PZ = value;
                    RaisePropertyChanged("PZ", true);
                }

            }
        }
        public System.DateTime SJ
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
        public System.Nullable<double> C
        {
            get
            {
                return this._C;
            }
            set
            {
                if (!_C.Equals(value))
                {
                    _C = value;
                    RaisePropertyChanged("C", true);
                }

            }
        }
        public System.Nullable<double> HUIFEN
        {
            get
            {
                return this._HUIFEN;
            }
            set
            {
                if (!_HUIFEN.Equals(value))
                {
                    _HUIFEN = value;
                    RaisePropertyChanged("HUIFEN", true);
                }

            }
        }
        public System.Nullable<double> HUIFA
        {
            get
            {
                return this._HUIFA;
            }
            set
            {
                if (!_HUIFA.Equals(value))
                {
                    _HUIFA = value;
                    RaisePropertyChanged("HUIFA", true);
                }

            }
        }
        public System.Nullable<double> S
        {
            get
            {
                return this._S;
            }
            set
            {
                if (!_S.Equals(value))
                {
                    _S = value;
                    RaisePropertyChanged("S", true);
                }

            }
        }
        [DisplayName("水份")]
        public System.Nullable<double> SF
        {
            get
            {
                return this._SF;
            }
            set
            {
                if (!_SF.Equals(value))
                {
                    _SF = value;
                    RaisePropertyChanged("SF", true);
                }

            }
        }

        [DisplayName("细度")]
        public System.Nullable<double> XIDU
        {
            get
            {
                return this._XIDU;
            }
            set
            {
                if (!_XIDU.Equals(value))
                {
                    _XIDU = value;
                    RaisePropertyChanged("XIDU", true);
                }

            }
        }

        [DisplayName("发热值")]
        public System.Nullable<double> FAREZHI
        {
            get
            {
                return this._FAREZHI;
            }
            set
            {
                if (!_FAREZHI.Equals(value))
                {
                    _FAREZHI = value;
                    RaisePropertyChanged("FAREZHI", true);
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
                insertCmd.CommandText = "INSERT INTO DDMF(PZ,SJ,C,HUIFEN,HUIFA,S,SF,FAREZHI,XIDU) VALUES(:PZ,:SJ,:C,:HUIFEN,:HUIFA,:S,:SF,:FAREZHI,:XIDU)";
                insertCmd.Parameters.Add(":PZ", OracleType.VarChar, 10).Value = this.PZ;
                insertCmd.Parameters.Add(":SJ", OracleType.DateTime, 7).Value = this.SJ;
                insertCmd.Parameters.Add(":C", OracleType.Double, 22).Value = (object)this.C ?? DBNull.Value;
                insertCmd.Parameters.Add(":HUIFEN", OracleType.Double, 22).Value = (object)this.HUIFEN ?? DBNull.Value;
                insertCmd.Parameters.Add(":HUIFA", OracleType.Double, 22).Value = (object)this.HUIFA ?? DBNull.Value;
                insertCmd.Parameters.Add(":S", OracleType.Double, 22).Value = (object)this.S ?? DBNull.Value;
                insertCmd.Parameters.Add(":SF", OracleType.Double, 22).Value = (object)this.SF ?? DBNull.Value;
                insertCmd.Parameters.Add(":FAREZHI", OracleType.Double, 22).Value = (object)this.FAREZHI ?? DBNull.Value;
                insertCmd.Parameters.Add(":XIDU", OracleType.Double, 22).Value = (object)this.XIDU ?? DBNull.Value;

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
                deleteCmd.CommandText = "DELETE FROM DDMF WHERE ROWID='" + RowId.Value + "'";

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
                updateCmd.CommandText = "UPDATE DDMF SET PZ=:PZ,SJ=:SJ,C=:C,HUIFEN=:HUIFEN,HUIFA=:HUIFA,S=:S,SF=:SF,FAREZHI=:FAREZHI,XIDU=:XIDU WHERE ROWID='" + RowId.Value + "'";
                updateCmd.Parameters.Add(":PZ", OracleType.VarChar, 10).Value = this.PZ;
                updateCmd.Parameters.Add(":SJ", OracleType.DateTime, 7).Value = this.SJ;
                updateCmd.Parameters.Add(":C", OracleType.Double, 22).Value = (object)this.C ?? DBNull.Value;
                updateCmd.Parameters.Add(":HUIFEN", OracleType.Double, 22).Value = (object)this.HUIFEN ?? DBNull.Value;
                updateCmd.Parameters.Add(":HUIFA", OracleType.Double, 22).Value = (object)this.HUIFA ?? DBNull.Value;
                updateCmd.Parameters.Add(":S", OracleType.Double, 22).Value = (object)this.S ?? DBNull.Value;
                updateCmd.Parameters.Add(":SF", OracleType.Double, 22).Value = (object)this.SF ?? DBNull.Value;
                updateCmd.Parameters.Add(":FAREZHI", OracleType.Double, 22).Value = (object)this.FAREZHI ?? DBNull.Value;
                updateCmd.Parameters.Add(":XIDU", OracleType.Double, 22).Value = (object)this.XIDU ?? DBNull.Value;

                updateCmd.ExecuteNonQuery();
            }


        }
    }
    public partial class DdmfTable : DbEntityTable<Ddmf>, IFilter
    {
        //Ddmf,ltznConnectionString
        public void LoadByRQ(DateTime rq)
        {
            //ConStr, EntiyType, SelectSql, SelectPara, SelectParaVal
            beginAdjustData = true;
            this.Clear();
            hideList.Clear();
            OracleConnection con = new OracleConnection(LTZN.Properties.Settings.Default.ltznConnectionString);
            con.Open();
            OracleCommand selectCmd = new OracleCommand();
            selectCmd.Connection = con;
            selectCmd.CommandText = "SELECT PZ,SJ,C,HUIFEN,HUIFA,S,SF,FAREZHI,XIDU,ROWID FROM DDMF WHERE TRUNC(SJ)=:RQ";
            selectCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = rq;

            OracleDataReader dr = selectCmd.ExecuteReader();
            while (dr.Read())
            {
                Ddmf item = new Ddmf();
                item.BeginInit();
                if (dr.IsDBNull(0)) item.PZ = ""; else item.PZ = dr.GetString(0);
                if (dr.IsDBNull(1)) item.SJ = DateTime.Now; else item.SJ = dr.GetDateTime(1);
                if (dr.IsDBNull(2)) item.C = null; else item.C = dr.GetDouble(2);
                if (dr.IsDBNull(3)) item.HUIFEN = null; else item.HUIFEN = dr.GetDouble(3);
                if (dr.IsDBNull(4)) item.HUIFA = null; else item.HUIFA = dr.GetDouble(4);
                if (dr.IsDBNull(5)) item.S = null; else item.S = dr.GetDouble(5);
                if (dr.IsDBNull(6)) item.SF = null; else item.SF = dr.GetDouble(6);
                if (dr.IsDBNull(7)) item.FAREZHI = null; else item.FAREZHI = dr.GetDouble(7);
                if (dr.IsDBNull(8)) item.XIDU = null; else item.XIDU = dr.GetDouble(8);
                item.RowId = dr.GetString(9);

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
    }
}
