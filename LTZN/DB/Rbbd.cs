using System;
using System.ComponentModel;
using System.Data.OracleClient;

namespace ZHCDB
{
    public partial class Rbbd : DbEntity
    {
        private string _BD = "";
        private System.Nullable<double> _SL;
        private string _DW = "";
        private string _FL = "";
        private string _SM = "";
        private System.Nullable<int> _GAOLU;
        private System.Nullable<System.DateTime> _RQ = null;
        private System.Nullable<System.DateTime> _SJ = null;
        //ConStr 

        [Query]
        [DisplayName("高炉")]
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
        public System.Nullable<System.DateTime> RQ
        {
            get
            {
                return this._RQ;
            }
            set
            {
                if (!_RQ.Equals(value))
                {
                    _RQ = value;
                    RaisePropertyChanged("RQ", true);
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
        [DisplayName("变动情况")]
        public string BD
        {
            get
            {
                return this._BD;
            }
            set
            {
                if (!_BD.Equals(value))
                {
                    _BD = value;
                    RaisePropertyChanged("BD", true);
                }

            }
        }

        [Query]
        [DisplayName("变动数量")]
        public System.Nullable<double> SL
        {
            get
            {
                return this._SL;
            }
            set
            {
                if (!_SL.Equals(value))
                {
                    _SL = value;
                    RaisePropertyChanged("SL", true);
                }

            }
        }

        [Query]
        [DisplayName("变动说明")]
        public string DW
        {
            get
            {
                return this._DW;
            }
            set
            {
                if (!_DW.Equals(value))
                {
                    _DW = value;
                    RaisePropertyChanged("DW", true);
                }

            }
        }

        [Query]
        [DisplayName("变动分类")]
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

        [Query]
        [DisplayName("变动说明")]
        public string SM
        {
            get
            {
                return this._SM;
            }
            set
            {
                if (!_SM.Equals(value))
                {
                    _SM = value;
                    RaisePropertyChanged("SM", true);
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
                insertCmd.CommandText = "INSERT INTO RBBD(BD,SL,DW,FL,SM,GAOLU,RQ,SJ) VALUES(:BD,:SL,:DW,:FL,:SM,:GAOLU,:RQ,:SJ)";
                insertCmd.Parameters.Add(":BD", OracleType.VarChar, 20).Value = this.BD;
                insertCmd.Parameters.Add(":SL", OracleType.Double, 22).Value = (object)this.SL ?? DBNull.Value;
                insertCmd.Parameters.Add(":DW", OracleType.VarChar, 20).Value = this.DW;
                insertCmd.Parameters.Add(":FL", OracleType.VarChar, 200).Value = this.FL;
                insertCmd.Parameters.Add(":SM", OracleType.VarChar, 1000).Value = this.SM;
                insertCmd.Parameters.Add(":GAOLU", OracleType.Int32, 22).Value = (object)this.GAOLU ?? DBNull.Value;
                insertCmd.Parameters.Add(":RQ", OracleType.DateTime, 7).Value = (object)this.RQ ?? DBNull.Value;
                insertCmd.Parameters.Add(":SJ", OracleType.DateTime, 7).Value = (object)this.SJ ?? DBNull.Value;

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
                deleteCmd.CommandText = "DELETE FROM RBBD WHERE ROWID='" + RowId.Value + "'";

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
                updateCmd.CommandText = "UPDATE RBBD SET BD=:BD,SL=:SL,DW=:DW,FL=:FL,SM=:SM,GAOLU=:GAOLU,RQ=:RQ,SJ=:SJ WHERE ROWID='" + RowId.Value + "'";
                updateCmd.Parameters.Add(":BD", OracleType.VarChar, 20).Value = this.BD;
                updateCmd.Parameters.Add(":SL", OracleType.Double, 22).Value = (object)this.SL ?? DBNull.Value;
                updateCmd.Parameters.Add(":DW", OracleType.VarChar, 20).Value = this.DW;
                updateCmd.Parameters.Add(":FL", OracleType.VarChar, 200).Value = this.FL;
                updateCmd.Parameters.Add(":SM", OracleType.VarChar, 1000).Value = this.SM;
                updateCmd.Parameters.Add(":GAOLU", OracleType.Int32, 22).Value = (object)this.GAOLU ?? DBNull.Value;
                updateCmd.Parameters.Add(":RQ", OracleType.DateTime, 7).Value = (object)this.RQ ?? DBNull.Value;
                updateCmd.Parameters.Add(":SJ", OracleType.DateTime, 7).Value = (object)this.SJ ?? DBNull.Value;

                updateCmd.ExecuteNonQuery();
            }


        }
    }
    public partial class RbbdTable : DbEntityTable<Rbbd>, IFilter
    {

        private DateTime? loadRq = null;
        private int? loadGaolu = null;

        //Rbbd,ltznConnectionString
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
            selectCmd.CommandText = "SELECT BD,SL,DW,FL,SM,GAOLU,RQ,SJ,ROWID FROM RBBD WHERE TRUNC(SJ)=:RQ AND GAOLU=:GAOLU order by SJ";
            selectCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = rq;
            selectCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = gaolu;

            OracleDataReader dr = selectCmd.ExecuteReader();
            while (dr.Read())
            {
                Rbbd item = new Rbbd();
                item.BeginInit();
                if (dr.IsDBNull(0)) item.BD = ""; else item.BD = dr.GetString(0);
                if (dr.IsDBNull(1)) item.SL = null; else item.SL = dr.GetDouble(1);
                if (dr.IsDBNull(2)) item.DW = ""; else item.DW = dr.GetString(2);
                if (dr.IsDBNull(3)) item.FL = ""; else item.FL = dr.GetString(3);
                if (dr.IsDBNull(4)) item.SM = ""; else item.SM = dr.GetString(4);
                if (dr.IsDBNull(5)) item.GAOLU = null; else item.GAOLU = dr.GetInt32(5);
                if (dr.IsDBNull(6)) item.RQ = null; else item.RQ = dr.GetDateTime(6);
                if (dr.IsDBNull(7)) item.SJ = null; else item.SJ = dr.GetDateTime(7);
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

        public void Load(string whereSql)
        {
            beginAdjustData = true;
            this.Clear();
            hideList.Clear();
            OracleConnection con = new OracleConnection(LTZN.Properties.Settings.Default.ltznConnectionString);
            con.Open();
            OracleCommand selectCmd = new OracleCommand();
            selectCmd.Connection = con;
            selectCmd.CommandText = "SELECT BD,SL,DW,FL,SM,GAOLU,RQ,SJ,ROWID FROM RBBD WHERE " + whereSql +" order by gaolu,sj";

            OracleDataReader dr = selectCmd.ExecuteReader();
            while (dr.Read())
            {
                Rbbd item = new Rbbd();
                item.BeginInit();
                if (dr.IsDBNull(0)) item.BD = ""; else item.BD = dr.GetString(0);
                if (dr.IsDBNull(1)) item.SL = null; else item.SL = dr.GetDouble(1);
                if (dr.IsDBNull(2)) item.DW = ""; else item.DW = dr.GetString(2);
                if (dr.IsDBNull(3)) item.FL = ""; else item.FL = dr.GetString(3);
                if (dr.IsDBNull(4)) item.SM = ""; else item.SM = dr.GetString(4);
                if (dr.IsDBNull(5)) item.GAOLU = null; else item.GAOLU = dr.GetInt32(5);
                if (dr.IsDBNull(6)) item.RQ = null; else item.RQ = dr.GetDateTime(6);
                if (dr.IsDBNull(7)) item.SJ = null; else item.SJ = dr.GetDateTime(7);
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

                    if (item.RQ == null)
                        item.RQ = new DateTime(loadRq.Value.Year, loadRq.Value.Month, loadRq.Value.Day,DateTime.Now.Hour,DateTime.Now.Minute, 0);

                    item.RQ = new DateTime(loadRq.Value.Year, loadRq.Value.Month, loadRq.Value.Day, item.RQ.Value.Hour, item.RQ.Value.Minute, 0);

                    item.GAOLU = loadGaolu.Value;
                    item.SJ = item.RQ;
                }
            }
            base.Save();
        }
    }
}


