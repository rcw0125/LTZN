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
    public partial class Rbgz : DbEntity
    {
        private DateTime _SJ = DateTime.Today;
        private string _BC = "";
        private string _JL = "";
        private int _GAOLU;
        //ConStr 



        public DateTime SJ
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
        public string JL
        {
            get
            {
                return this._JL;
            }
            set
            {
                if (!_JL.Equals(value))
                {
                    _JL = value;
                    RaisePropertyChanged("JL", true);
                }

            }
        }
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

        public bool CheckPk(OracleTransaction trans)
        {
            OracleConnection con = trans.Connection;
            OracleCommand checkCmd = new OracleCommand();
            checkCmd.Connection = con;
            checkCmd.Transaction = trans;
            checkCmd.CommandText = "SELECT ROWID FROM RBGZ WHERE TRUNC(SJ)=:RQ AND GAOLU=:GAOLU AND BC=:BC";
            checkCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = this.SJ;
            checkCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = this.GAOLU;
            checkCmd.Parameters.Add(":BC", OracleType.VarChar, 6).Value = this.BC;
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
                if (CheckPk(trans))
                {
                    DataState = DataStateType.Update;
                    DbUpdate(trans);
                    return;
                }
                OracleConnection con = trans.Connection;
                OracleCommand insertCmd = new OracleCommand();
                insertCmd.Connection = con;
                insertCmd.Transaction = trans;
                insertCmd.CommandText = "INSERT INTO RBGZ(SJ,BC,JL,GAOLU) VALUES(:SJ,:BC,:JL,:GAOLU)";
                insertCmd.Parameters.Add(":SJ", OracleType.DateTime, 7).Value = (object)this.SJ ?? DBNull.Value;
                insertCmd.Parameters.Add(":BC", OracleType.VarChar, 6).Value = this.BC;
                insertCmd.Parameters.Add(":JL", OracleType.VarChar, 2000).Value = this.JL;
                insertCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = this.GAOLU;

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
                deleteCmd.CommandText = "DELETE FROM RBGZ WHERE ROWID='" + RowId.Value + "'";

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
                updateCmd.CommandText = "UPDATE RBGZ SET SJ=:SJ,BC=:BC,JL=:JL,GAOLU=:GAOLU WHERE ROWID='" + RowId.Value + "'";
                updateCmd.Parameters.Add(":SJ", OracleType.DateTime, 7).Value = (object)this.SJ ?? DBNull.Value;
                updateCmd.Parameters.Add(":BC", OracleType.VarChar, 6).Value = this.BC;
                updateCmd.Parameters.Add(":JL", OracleType.VarChar, 2000).Value = this.JL;
                updateCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = this.GAOLU;

                updateCmd.ExecuteNonQuery();
            }


        }
    }
    public partial class RbgzTable : DbEntityTable<Rbgz>, IFilter
    {
        //Rbgz,ltznConnectionString
        public void Load(int gaolu, DateTime rq)
        {
            //ConStr, EntiyType, SelectSql, SelectPara, SelectParaVal
            beginAdjustData = true;
            this.Clear();
            hideList.Clear();
            OracleConnection con = new OracleConnection(LTZN.Properties.Settings.Default.ltznConnectionString);
            con.Open();
            OracleCommand selectCmd = new OracleCommand();
            selectCmd.Connection = con;
            selectCmd.CommandText = "SELECT SJ,BC,JL,GAOLU,ROWID FROM RBGZ WHERE TRUNC(SJ)=:RQ AND GAOLU=:GAOLU";
            selectCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = rq;
            selectCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = gaolu;

            OracleDataReader dr = selectCmd.ExecuteReader();
            while (dr.Read())
            {
                Rbgz item = new Rbgz();
                item.BeginInit();
                if (dr.IsDBNull(0)) item.SJ = DateTime.Today; else item.SJ = dr.GetDateTime(0);
                if (dr.IsDBNull(1)) item.BC = ""; else item.BC = dr.GetString(1);
                if (dr.IsDBNull(2)) item.JL = ""; else item.JL = dr.GetString(2);
                if (dr.IsDBNull(3)) item.GAOLU = 0; else item.GAOLU = dr.GetInt32(3);
                item.RowId = dr.GetString(4);

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
