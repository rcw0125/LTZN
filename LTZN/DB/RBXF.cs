using System;
using System.ComponentModel;
using System.Data.OracleClient;

namespace ZHCDB
{
    public partial class Rbxf : DbEntity
    {
        private System.Nullable<int> _GAOLU;
        private string _MC = "";
        private System.Nullable<System.DateTime> _KS = null;
        private System.Nullable<System.DateTime> _ZZ = null;
        private string _FL = "";
        private string _YY = "";
        private System.Nullable<System.DateTime> _SJ = null;
        private System.Nullable<double> _JG;
        //ConStr 



        public System.Nullable<int> GAOLU
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
        public string MC
        {
            get
            {
                return this._MC;
            }
            set
            {
                if (!_MC.Equals(value))
                {
                    _MC = value;
                    RaisePropertyChanged("MC", true);
                }

            }
        }
        public System.Nullable<System.DateTime> KS
        {
            get
            {
                return this._KS;
            }
            set
            {
                if (!_KS.Equals(value))
                {
                    _KS = value;
                    RaisePropertyChanged("KS", true);
                    if (KS != null && ZZ != null)
                    {
                        JG = (ZZ.Value - KS.Value).TotalMinutes;
                    }
                    else
                    {
                        JG = null;
                    }
                }
            }
        }
        public System.Nullable<System.DateTime> ZZ
        {
            get
            {
                return this._ZZ;
            }
            set
            {
                if (!_ZZ.Equals(value))
                {
                    _ZZ = value;
                    RaisePropertyChanged("ZZ", true);

                    if (KS != null && ZZ != null)
                    {
                        JG = (ZZ.Value - KS.Value).TotalMinutes;
                    }
                    else
                    {
                        JG = null;
                    }
                }

            }
        }
        public string FL
        {
            get
            {
                return this._FL;
            }
            set
            {
                if (!_FL.Equals(value))
                {
                    _FL = value;
                    RaisePropertyChanged("FL", true);
                }

            }
        }
        public string YY
        {
            get
            {
                return this._YY;
            }
            set
            {
                if (!_YY.Equals(value))
                {
                    _YY = value;
                    RaisePropertyChanged("YY", true);
                }

            }
        }
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
                    if (_SJ != null)
                    {
                        if (KS != null)
                        {
                            KS = new DateTime(_SJ.Value.Year, _SJ.Value.Month, _SJ.Value.Day, KS.Value.Hour, KS.Value.Minute, 0);
                        }
                        if (ZZ != null)
                        {
                            ZZ = new DateTime(_SJ.Value.Year, _SJ.Value.Month, _SJ.Value.Day, ZZ.Value.Hour, ZZ.Value.Minute, 0);
                        }
                    }
                }

            }
        }
        public System.Nullable<double> JG
        {
            get
            {
                return this._JG;
            }
            set
            {
                if (!_JG.Equals(value))
                {
                    _JG = value;
                    RaisePropertyChanged("JG", true);
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
                insertCmd.CommandText = "INSERT INTO RBXF(GAOLU,MC,KS,ZZ,FL,YY,SJ,JG) VALUES(:GAOLU,:MC,:KS,:ZZ,:FL,:YY,:SJ,:JG)";
                insertCmd.Parameters.Add(":GAOLU", OracleType.Int32, 22).Value = (object)this.GAOLU ?? DBNull.Value;
                insertCmd.Parameters.Add(":MC", OracleType.VarChar, 20).Value = this.MC;
                insertCmd.Parameters.Add(":KS", OracleType.DateTime, 7).Value = (object)this.KS ?? DBNull.Value;
                insertCmd.Parameters.Add(":ZZ", OracleType.DateTime, 7).Value = (object)this.ZZ ?? DBNull.Value;
                insertCmd.Parameters.Add(":FL", OracleType.VarChar, 2000).Value = this.FL;
                insertCmd.Parameters.Add(":YY", OracleType.VarChar, 2000).Value = this.YY;
                insertCmd.Parameters.Add(":SJ", OracleType.DateTime, 7).Value = (object)this.SJ ?? DBNull.Value;
                insertCmd.Parameters.Add(":JG", OracleType.Double, 22).Value = (object)this.JG ?? DBNull.Value;

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
                deleteCmd.CommandText = "DELETE FROM RBXF WHERE ROWID='" + RowId.Value + "'";

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
                updateCmd.CommandText = "UPDATE RBXF SET GAOLU=:GAOLU,MC=:MC,KS=:KS,ZZ=:ZZ,FL=:FL,YY=:YY,SJ=:SJ,JG=:JG WHERE ROWID='" + RowId.Value + "'";
                updateCmd.Parameters.Add(":GAOLU", OracleType.Int32, 22).Value = (object)this.GAOLU ?? DBNull.Value;
                updateCmd.Parameters.Add(":MC", OracleType.VarChar, 20).Value = this.MC;
                updateCmd.Parameters.Add(":KS", OracleType.DateTime, 7).Value = (object)this.KS ?? DBNull.Value;
                updateCmd.Parameters.Add(":ZZ", OracleType.DateTime, 7).Value = (object)this.ZZ ?? DBNull.Value;
                updateCmd.Parameters.Add(":FL", OracleType.VarChar, 2000).Value = this.FL;
                updateCmd.Parameters.Add(":YY", OracleType.VarChar, 2000).Value = this.YY;
                updateCmd.Parameters.Add(":SJ", OracleType.DateTime, 7).Value = (object)this.SJ ?? DBNull.Value;
                updateCmd.Parameters.Add(":JG", OracleType.Double, 22).Value = (object)this.JG ?? DBNull.Value;

                updateCmd.ExecuteNonQuery();
            }


        }
    }
    public partial class RbxfTable : DbEntityTable<Rbxf>, IFilter
    {

        private DateTime? loadRq = null;
        private int? loadGaolu = null;

        //Rbxf,ltznConnectionString
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
            selectCmd.CommandText = "SELECT GAOLU,MC,KS,ZZ,FL,YY,SJ,JG,ROWID FROM RBXF WHERE TRUNC(SJ)=:RQ AND GAOLU=:GAOLU order by KS";
            selectCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = rq;
            selectCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = gaolu;

            OracleDataReader dr = selectCmd.ExecuteReader();
            while (dr.Read())
            {
                Rbxf item = new Rbxf();
                item.BeginInit();
                if (dr.IsDBNull(0)) item.GAOLU = null; else item.GAOLU = dr.GetInt32(0);
                if (dr.IsDBNull(1)) item.MC = ""; else item.MC = dr.GetString(1);
                if (dr.IsDBNull(2)) item.KS = null; else item.KS = dr.GetDateTime(2);
                if (dr.IsDBNull(3)) item.ZZ = null; else item.ZZ = dr.GetDateTime(3);
                if (dr.IsDBNull(4)) item.FL = ""; else item.FL = dr.GetString(4);
                if (dr.IsDBNull(5)) item.YY = ""; else item.YY = dr.GetString(5);
                if (dr.IsDBNull(6)) item.SJ = null; else item.SJ = dr.GetDateTime(6);
                if (dr.IsDBNull(7)) item.JG = null; else item.JG = dr.GetDouble(7);
                item.RowId = dr.GetString(8);

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
                foreach (var item in this)
                {
                    item.SJ = loadRq.Value;
                    item.GAOLU = loadGaolu.Value;
                }
            }
            base.Save();
        }
    }
}
