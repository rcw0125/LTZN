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
    public partial class Rbliaoxian : DbEntity
    {
        private System.Nullable<double> _LIAOQIAN;
        private System.Nullable<double> _LIAOHOU;
        private System.Nullable<int> _PISHU;
        private System.DateTime _T = DateTime.Now;
        private System.Nullable<int> _GAOLU;
        private System.Nullable<double> _JIANGE;
        private string _JIAJIAO = "";
        private string _BC = "";
        private string _SJ = "";
        private System.Nullable<double> _XIUGAI;
        private System.Nullable<double> _JISHAO;
        private System.Nullable<double> _QIUTUAN;
        private System.Nullable<double> _JIAOTAN;
        private System.Nullable<double> _JIAODING;
        private System.Nullable<double> _KUAIKUANG;
        private System.Nullable<double> _QITAKUANG;
        //ConStr 



        public System.Nullable<double> LIAOQIAN
        {
            get
            {
                return this._LIAOQIAN;
            }
            set
            {
                if (!_LIAOQIAN.Equals(value))
                {
                    _LIAOQIAN = value;
                    RaisePropertyChanged("LIAOQIAN", true);
                }

            }
        }
        public System.Nullable<double> LIAOHOU
        {
            get
            {
                return this._LIAOHOU;
            }
            set
            {
                if (!_LIAOHOU.Equals(value))
                {
                    _LIAOHOU = value;
                    RaisePropertyChanged("LIAOHOU", true);
                }

            }
        }
        public System.Nullable<int> PISHU
        {
            get
            {
                return this._PISHU;
            }
            set
            {
                if (!_PISHU.Equals(value))
                {
                    _PISHU = value;
                    RaisePropertyChanged("PISHU", true);
                }

            }
        }
        public System.DateTime T
        {
            get
            {
                return this._T;
            }
            set
            {
                if (!_T.Equals(value))
                {
                    _T = value;
                    RaisePropertyChanged("T", true);
                }

            }
        }
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
        public System.Nullable<double> JIANGE
        {
            get
            {
                return this._JIANGE;
            }
            set
            {
                if (!_JIANGE.Equals(value))
                {
                    _JIANGE = value;
                    RaisePropertyChanged("JIANGE", true);
                }

            }
        }
        public string JIAJIAO
        {
            get
            {
                return this._JIAJIAO;
            }
            set
            {
                if (!_JIAJIAO.Equals(value))
                {
                    _JIAJIAO = value;
                    RaisePropertyChanged("JIAJIAO", true);
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
        public string SJ
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
        public System.Nullable<double> XIUGAI
        {
            get
            {
                return this._XIUGAI;
            }
            set
            {
                if (!_XIUGAI.Equals(value))
                {
                    _XIUGAI = value;
                    RaisePropertyChanged("XIUGAI", true);
                }

            }
        }
        public System.Nullable<double> JISHAO
        {
            get
            {
                return this._JISHAO;
            }
            set
            {
                if (!_JISHAO.Equals(value))
                {
                    _JISHAO = value;
                    RaisePropertyChanged("JISHAO", true);
                }

            }
        }
        public System.Nullable<double> QIUTUAN
        {
            get
            {
                return this._QIUTUAN;
            }
            set
            {
                if (!_QIUTUAN.Equals(value))
                {
                    _QIUTUAN = value;
                    RaisePropertyChanged("QIUTUAN", true);
                }

            }
        }
        public System.Nullable<double> JIAOTAN
        {
            get
            {
                return this._JIAOTAN;
            }
            set
            {
                if (!_JIAOTAN.Equals(value))
                {
                    _JIAOTAN = value;
                    RaisePropertyChanged("JIAOTAN", true);
                }

            }
        }
        public System.Nullable<double> JIAODING
        {
            get
            {
                return this._JIAODING;
            }
            set
            {
                if (!_JIAODING.Equals(value))
                {
                    _JIAODING = value;
                    RaisePropertyChanged("JIAODING", true);
                }

            }
        }
        public System.Nullable<double> KUAIKUANG
        {
            get
            {
                return this._KUAIKUANG;
            }
            set
            {
                if (!_KUAIKUANG.Equals(value))
                {
                    _KUAIKUANG = value;
                    RaisePropertyChanged("KUAIKUANG", true);
                }

            }
        }
        public System.Nullable<double> QITAKUANG
        {
            get
            {
                return this._QITAKUANG;
            }
            set
            {
                if (!_QITAKUANG.Equals(value))
                {
                    _QITAKUANG = value;
                    RaisePropertyChanged("QITAKUANG", true);
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
                insertCmd.CommandText = "INSERT INTO RBLIAOXIAN(LIAOQIAN,LIAOHOU,PISHU,T,GAOLU,JIANGE,JIAJIAO,BC,SJ,XIUGAI,JISHAO,QIUTUAN,JIAOTAN,JIAODING,KUAIKUANG,QITAKUANG) VALUES(:LIAOQIAN,:LIAOHOU,:PISHU,:T,:GAOLU,:JIANGE,:JIAJIAO,:BC,:SJ,:XIUGAI,:JISHAO,:QIUTUAN,:JIAOTAN,:JIAODING,:KUAIKUANG,:QITAKUANG)";
                insertCmd.Parameters.Add(":LIAOQIAN", OracleType.Double).Value = (object)this.LIAOQIAN ?? DBNull.Value;
                insertCmd.Parameters.Add(":LIAOHOU", OracleType.Double).Value = (object)this.LIAOHOU ?? DBNull.Value;
                insertCmd.Parameters.Add(":PISHU", OracleType.Int32, 22).Value = (object)this.PISHU ?? DBNull.Value;
                insertCmd.Parameters.Add(":T", OracleType.DateTime, 7).Value = this.T;
                insertCmd.Parameters.Add(":GAOLU", OracleType.Int32, 22).Value = (object)this.GAOLU ?? DBNull.Value;
                insertCmd.Parameters.Add(":JIANGE", OracleType.Double).Value = (object)this.JIANGE ?? DBNull.Value;
                insertCmd.Parameters.Add(":JIAJIAO", OracleType.VarChar, 20).Value = this.JIAJIAO ?? "";
                insertCmd.Parameters.Add(":BC", OracleType.VarChar, 6).Value = this.BC??"";
                insertCmd.Parameters.Add(":SJ", OracleType.VarChar, 8).Value = this.SJ??"";
                insertCmd.Parameters.Add(":XIUGAI", OracleType.Double).Value = (object)this.XIUGAI ?? DBNull.Value;
                insertCmd.Parameters.Add(":JISHAO", OracleType.Double).Value = (object)this.JISHAO ?? DBNull.Value;
                insertCmd.Parameters.Add(":QIUTUAN", OracleType.Double).Value = (object)this.QIUTUAN ?? DBNull.Value;
                insertCmd.Parameters.Add(":JIAOTAN", OracleType.Double).Value = (object)this.JIAOTAN ?? DBNull.Value;
                insertCmd.Parameters.Add(":JIAODING", OracleType.Double).Value = (object)this.JIAODING ?? DBNull.Value;
                insertCmd.Parameters.Add(":KUAIKUANG", OracleType.Double).Value = (object)this.KUAIKUANG ?? DBNull.Value;
                insertCmd.Parameters.Add(":QITAKUANG", OracleType.Double).Value = (object)this.QITAKUANG ?? DBNull.Value;

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
                deleteCmd.CommandText = "DELETE FROM RBLIAOXIAN WHERE ROWID='" + RowId.Value + "'";

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
                updateCmd.CommandText = "UPDATE RBLIAOXIAN SET LIAOQIAN=:LIAOQIAN,LIAOHOU=:LIAOHOU,PISHU=:PISHU,T=:T,GAOLU=:GAOLU,JIANGE=:JIANGE,JIAJIAO=:JIAJIAO,BC=:BC,SJ=:SJ,XIUGAI=:XIUGAI,JISHAO=:JISHAO,QIUTUAN=:QIUTUAN,JIAOTAN=:JIAOTAN,JIAODING=:JIAODING,KUAIKUANG=:KUAIKUANG,QITAKUANG=:QITAKUANG WHERE ROWID='" + RowId.Value + "'";
                updateCmd.Parameters.Add(":LIAOQIAN", OracleType.Double).Value = (object)this.LIAOQIAN ?? DBNull.Value;
                updateCmd.Parameters.Add(":LIAOHOU", OracleType.Double).Value = (object)this.LIAOHOU ?? DBNull.Value;
                updateCmd.Parameters.Add(":PISHU", OracleType.Int32, 22).Value = (object)this.PISHU ?? DBNull.Value;
                updateCmd.Parameters.Add(":T", OracleType.DateTime, 7).Value = this.T;
                updateCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = (object)this.GAOLU ?? DBNull.Value;
                updateCmd.Parameters.Add(":JIANGE", OracleType.Double).Value = (object)this.JIANGE ?? DBNull.Value;
                updateCmd.Parameters.Add(":JIAJIAO", OracleType.VarChar, 20).Value = this.JIAJIAO ?? "";
                updateCmd.Parameters.Add(":BC", OracleType.VarChar, 6).Value = this.BC??"";
                updateCmd.Parameters.Add(":SJ", OracleType.VarChar, 8).Value = this.SJ ?? "";
                updateCmd.Parameters.Add(":XIUGAI", OracleType.Double).Value = (object)this.XIUGAI ?? DBNull.Value;
                updateCmd.Parameters.Add(":JISHAO", OracleType.Double).Value = (object)this.JISHAO ?? DBNull.Value;
                updateCmd.Parameters.Add(":QIUTUAN", OracleType.Double).Value = (object)this.QIUTUAN ?? DBNull.Value;
                updateCmd.Parameters.Add(":JIAOTAN", OracleType.Double).Value = (object)this.JIAOTAN ?? DBNull.Value;
                updateCmd.Parameters.Add(":JIAODING", OracleType.Double).Value = (object)this.JIAODING ?? DBNull.Value;
                updateCmd.Parameters.Add(":KUAIKUANG", OracleType.Double).Value = (object)this.KUAIKUANG ?? DBNull.Value;
                updateCmd.Parameters.Add(":QITAKUANG", OracleType.Double).Value = (object)this.QITAKUANG ?? DBNull.Value;

                updateCmd.ExecuteNonQuery();
            }


        }
    }
    public partial class RbliaoxianTable : DbEntityTable<Rbliaoxian>, IFilter
    {
        //Rbliaoxian,ltznConnectionString
        public void LoadByRQGaoluBc(DateTime rq,int gaolu,string bc)
        {
            //ConStr, EntiyType, SelectSql, SelectPara, SelectParaVal
            beginAdjustData = true;
            this.Clear();
            hideList.Clear();
            OracleConnection con = new OracleConnection(LTZN.Properties.Settings.Default.ltznConnectionString);
            con.Open();
            OracleCommand selectCmd = new OracleCommand();
            selectCmd.Connection = con;
            selectCmd.CommandText = "SELECT LIAOQIAN,LIAOHOU,PISHU,T,GAOLU,JIANGE,JIAJIAO,BC,SJ,XIUGAI,JISHAO,QIUTUAN,JIAOTAN,JIAODING,KUAIKUANG,QITAKUANG,ROWID FROM RBLIAOXIAN WHERE trunc(t)=:RQ and gaolu=:GAOLU and bc=:BC order by t";
            selectCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = rq;
            selectCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = gaolu;
            selectCmd.Parameters.Add(":BC", OracleType.VarChar).Value = bc;

            OracleDataReader dr = selectCmd.ExecuteReader();
            while (dr.Read())
            {
                Rbliaoxian item = new Rbliaoxian();
                item.BeginInit();
                if (dr.IsDBNull(0)) item.LIAOQIAN = null; else item.LIAOQIAN = dr.GetDouble(0);
                if (dr.IsDBNull(1)) item.LIAOHOU = null; else item.LIAOHOU = dr.GetDouble(1);
                if (dr.IsDBNull(2)) item.PISHU = null; else item.PISHU = dr.GetInt32(2);
                if (dr.IsDBNull(3)) item.T = DateTime.Now; else item.T = dr.GetDateTime(3);
                if (dr.IsDBNull(4)) item.GAOLU = null; else item.GAOLU = dr.GetInt32(4);
                if (dr.IsDBNull(5)) item.JIANGE = null; else item.JIANGE = dr.GetDouble(5);
                if (dr.IsDBNull(6)) item.JIAJIAO = ""; else item.JIAJIAO = dr.GetString(6);
                if (dr.IsDBNull(7)) item.BC = ""; else item.BC = dr.GetString(7);
                if (dr.IsDBNull(8)) item.SJ = ""; else item.SJ = dr.GetString(8);
                if (dr.IsDBNull(9)) item.XIUGAI = null; else item.XIUGAI = dr.GetDouble(9);
                if (dr.IsDBNull(10)) item.JISHAO = null; else item.JISHAO = dr.GetDouble(10);
                if (dr.IsDBNull(11)) item.QIUTUAN = null; else item.QIUTUAN = dr.GetDouble(11);
                if (dr.IsDBNull(12)) item.JIAOTAN = null; else item.JIAOTAN = dr.GetDouble(12);
                if (dr.IsDBNull(13)) item.JIAODING = null; else item.JIAODING = dr.GetDouble(13);
                if (dr.IsDBNull(14)) item.KUAIKUANG = null; else item.KUAIKUANG = dr.GetDouble(14);
                if (dr.IsDBNull(15)) item.QITAKUANG = null; else item.QITAKUANG = dr.GetDouble(15);
                item.RowId = dr.GetString(16);

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

   
  
