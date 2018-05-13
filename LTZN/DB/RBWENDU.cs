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
    public partial class Rbwendu : DbEntity
    {
        private int _GAOLU;
        private System.DateTime _SJ = DateTime.Now;
        private System.Nullable<double> _WD1;
        private System.Nullable<double> _WD2;
        private System.Nullable<double> _WD3;
        private System.Nullable<double> _WD4;
        private System.Nullable<double> _WD5;
        private System.Nullable<double> _WD6;
        private System.Nullable<double> _WD7;
        private System.Nullable<double> _WD8;
        private System.Nullable<double> _WD9;
        private System.Nullable<double> _WD10;
        private System.Nullable<double> _WD11;
        private System.Nullable<double> _WD12;
        private System.Nullable<double> _WD13;
        private System.Nullable<double> _WD14;
        private System.Nullable<double> _WD15;
        private System.Nullable<double> _WD16;
        private System.Nullable<double> _WD17;
        private System.Nullable<double> _WD18;
        private System.Nullable<double> _WD19;
        private System.Nullable<double> _WD20;
        private System.Nullable<double> _WD21;
        private System.Nullable<double> _WD22;
        private System.Nullable<double> _WD23;
        private System.Nullable<double> _WD24;
        private System.Nullable<double> _WD25;
        private System.Nullable<double> _WD26;
        private System.Nullable<double> _WD27;
        private System.Nullable<double> _WD28;
        private System.Nullable<double> _WD29;
        private System.Nullable<double> _WD30;
        private string _BZ = "";
        //ConStr 



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
        public System.Nullable<double> WD1
        {
            get
            {
                return this._WD1;
            }
            set
            {
                if (!_WD1.Equals(value))
                {
                    _WD1 = value;
                    RaisePropertyChanged("WD1", true);
                }

            }
        }
        public System.Nullable<double> WD2
        {
            get
            {
                return this._WD2;
            }
            set
            {
                if (!_WD2.Equals(value))
                {
                    _WD2 = value;
                    RaisePropertyChanged("WD2", true);
                }

            }
        }
        public System.Nullable<double> WD3
        {
            get
            {
                return this._WD3;
            }
            set
            {
                if (!_WD3.Equals(value))
                {
                    _WD3 = value;
                    RaisePropertyChanged("WD3", true);
                }

            }
        }
        public System.Nullable<double> WD4
        {
            get
            {
                return this._WD4;
            }
            set
            {
                if (!_WD4.Equals(value))
                {
                    _WD4 = value;
                    RaisePropertyChanged("WD4", true);
                }

            }
        }
        public System.Nullable<double> WD5
        {
            get
            {
                return this._WD5;
            }
            set
            {
                if (!_WD5.Equals(value))
                {
                    _WD5 = value;
                    RaisePropertyChanged("WD5", true);
                }

            }
        }
        public System.Nullable<double> WD6
        {
            get
            {
                return this._WD6;
            }
            set
            {
                if (!_WD6.Equals(value))
                {
                    _WD6 = value;
                    RaisePropertyChanged("WD6", true);
                }

            }
        }
        public System.Nullable<double> WD7
        {
            get
            {
                return this._WD7;
            }
            set
            {
                if (!_WD7.Equals(value))
                {
                    _WD7 = value;
                    RaisePropertyChanged("WD7", true);
                }

            }
        }
        public System.Nullable<double> WD8
        {
            get
            {
                return this._WD8;
            }
            set
            {
                if (!_WD8.Equals(value))
                {
                    _WD8 = value;
                    RaisePropertyChanged("WD8", true);
                }

            }
        }
        public System.Nullable<double> WD9
        {
            get
            {
                return this._WD9;
            }
            set
            {
                if (!_WD9.Equals(value))
                {
                    _WD9 = value;
                    RaisePropertyChanged("WD9", true);
                }

            }
        }
        public System.Nullable<double> WD10
        {
            get
            {
                return this._WD10;
            }
            set
            {
                if (!_WD10.Equals(value))
                {
                    _WD10 = value;
                    RaisePropertyChanged("WD10", true);
                }

            }
        }
        public System.Nullable<double> WD11
        {
            get
            {
                return this._WD11;
            }
            set
            {
                if (!_WD11.Equals(value))
                {
                    _WD11 = value;
                    RaisePropertyChanged("WD11", true);
                }

            }
        }
        public System.Nullable<double> WD12
        {
            get
            {
                return this._WD12;
            }
            set
            {
                if (!_WD12.Equals(value))
                {
                    _WD12 = value;
                    RaisePropertyChanged("WD12", true);
                }

            }
        }
        public System.Nullable<double> WD13
        {
            get
            {
                return this._WD13;
            }
            set
            {
                if (!_WD13.Equals(value))
                {
                    _WD13 = value;
                    RaisePropertyChanged("WD13", true);
                }

            }
        }
        public System.Nullable<double> WD14
        {
            get
            {
                return this._WD14;
            }
            set
            {
                if (!_WD14.Equals(value))
                {
                    _WD14 = value;
                    RaisePropertyChanged("WD14", true);
                }

            }
        }
        public System.Nullable<double> WD15
        {
            get
            {
                return this._WD15;
            }
            set
            {
                if (!_WD15.Equals(value))
                {
                    _WD15 = value;
                    RaisePropertyChanged("WD15", true);
                }

            }
        }
        public System.Nullable<double> WD16
        {
            get
            {
                return this._WD16;
            }
            set
            {
                if (!_WD16.Equals(value))
                {
                    _WD16 = value;
                    RaisePropertyChanged("WD16", true);
                }

            }
        }
        public System.Nullable<double> WD17
        {
            get
            {
                return this._WD17;
            }
            set
            {
                if (!_WD17.Equals(value))
                {
                    _WD17 = value;
                    RaisePropertyChanged("WD17", true);
                }

            }
        }
        public System.Nullable<double> WD18
        {
            get
            {
                return this._WD18;
            }
            set
            {
                if (!_WD18.Equals(value))
                {
                    _WD18 = value;
                    RaisePropertyChanged("WD18", true);
                }

            }
        }
        public System.Nullable<double> WD19
        {
            get
            {
                return this._WD19;
            }
            set
            {
                if (!_WD19.Equals(value))
                {
                    _WD19 = value;
                    RaisePropertyChanged("WD19", true);
                }

            }
        }
        public System.Nullable<double> WD20
        {
            get
            {
                return this._WD20;
            }
            set
            {
                if (!_WD20.Equals(value))
                {
                    _WD20 = value;
                    RaisePropertyChanged("WD20", true);
                }

            }
        }
        public System.Nullable<double> WD21
        {
            get
            {
                return this._WD21;
            }
            set
            {
                if (!_WD21.Equals(value))
                {
                    _WD21 = value;
                    RaisePropertyChanged("WD21", true);
                }

            }
        }
        public System.Nullable<double> WD22
        {
            get
            {
                return this._WD22;
            }
            set
            {
                if (!_WD22.Equals(value))
                {
                    _WD22 = value;
                    RaisePropertyChanged("WD22", true);
                }

            }
        }
        public System.Nullable<double> WD23
        {
            get
            {
                return this._WD23;
            }
            set
            {
                if (!_WD23.Equals(value))
                {
                    _WD23 = value;
                    RaisePropertyChanged("WD23", true);
                }

            }
        }
        public System.Nullable<double> WD24
        {
            get
            {
                return this._WD24;
            }
            set
            {
                if (!_WD24.Equals(value))
                {
                    _WD24 = value;
                    RaisePropertyChanged("WD24", true);
                }

            }
        }
        public System.Nullable<double> WD25
        {
            get
            {
                return this._WD25;
            }
            set
            {
                if (!_WD25.Equals(value))
                {
                    _WD25 = value;
                    RaisePropertyChanged("WD25", true);
                }

            }
        }
        public System.Nullable<double> WD26
        {
            get
            {
                return this._WD26;
            }
            set
            {
                if (!_WD26.Equals(value))
                {
                    _WD26 = value;
                    RaisePropertyChanged("WD26", true);
                }

            }
        }
        public System.Nullable<double> WD27
        {
            get
            {
                return this._WD27;
            }
            set
            {
                if (!_WD27.Equals(value))
                {
                    _WD27 = value;
                    RaisePropertyChanged("WD27", true);
                }

            }
        }
        public System.Nullable<double> WD28
        {
            get
            {
                return this._WD28;
            }
            set
            {
                if (!_WD28.Equals(value))
                {
                    _WD28 = value;
                    RaisePropertyChanged("WD28", true);
                }

            }
        }
        public System.Nullable<double> WD29
        {
            get
            {
                return this._WD29;
            }
            set
            {
                if (!_WD29.Equals(value))
                {
                    _WD29 = value;
                    RaisePropertyChanged("WD29", true);
                }

            }
        }
        public System.Nullable<double> WD30
        {
            get
            {
                return this._WD30;
            }
            set
            {
                if (!_WD30.Equals(value))
                {
                    _WD30 = value;
                    RaisePropertyChanged("WD30", true);
                }

            }
        }
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
    
        public virtual void DbInsert(OracleTransaction trans)
        {
            //InsertSql,InsertPara,AutoSql
            if (DataState == DataStateType.Add)
            {
                OracleConnection con = trans.Connection;
                OracleCommand insertCmd = new OracleCommand();
                insertCmd.Connection = con;
                insertCmd.Transaction = trans;
                insertCmd.CommandText = "INSERT INTO RBWENDU(GAOLU,SJ,WD1,WD2,WD3,WD4,WD5,WD6,WD7,WD8,WD9,WD10,WD11,WD12,WD13,WD14,WD15,WD16,WD17,WD18,WD19,WD20,WD21,WD22,WD23,WD24,WD25,WD26,WD27,WD28,WD29,WD30,BZ) VALUES(:GAOLU,:SJ,:WD1,:WD2,:WD3,:WD4,:WD5,:WD6,:WD7,:WD8,:WD9,:WD10,:WD11,:WD12,:WD13,:WD14,:WD15,:WD16,:WD17,:WD18,:WD19,:WD20,:WD21,:WD22,:WD23,:WD24,:WD25,:WD26,:WD27,:WD28,:WD29,:WD30,:BZ)";
                insertCmd.Parameters.Add(":GAOLU", OracleType.Int32, 22).Value = this.GAOLU;
                insertCmd.Parameters.Add(":SJ", OracleType.DateTime, 7).Value = this.SJ;
                insertCmd.Parameters.Add(":WD1", OracleType.Double, 22).Value = (object)this.WD1 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD2", OracleType.Double, 22).Value = (object)this.WD2 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD3", OracleType.Double, 22).Value = (object)this.WD3 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD4", OracleType.Double, 22).Value = (object)this.WD4 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD5", OracleType.Double, 22).Value = (object)this.WD5 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD6", OracleType.Double, 22).Value = (object)this.WD6 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD7", OracleType.Double, 22).Value = (object)this.WD7 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD8", OracleType.Double, 22).Value = (object)this.WD8 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD9", OracleType.Double, 22).Value = (object)this.WD9 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD10", OracleType.Double, 22).Value = (object)this.WD10 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD11", OracleType.Double, 22).Value = (object)this.WD11 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD12", OracleType.Double, 22).Value = (object)this.WD12 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD13", OracleType.Double, 22).Value = (object)this.WD13 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD14", OracleType.Double, 22).Value = (object)this.WD14 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD15", OracleType.Double, 22).Value = (object)this.WD15 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD16", OracleType.Double, 22).Value = (object)this.WD16 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD17", OracleType.Double, 22).Value = (object)this.WD17 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD18", OracleType.Double, 22).Value = (object)this.WD18 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD19", OracleType.Double, 22).Value = (object)this.WD19 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD20", OracleType.Double, 22).Value = (object)this.WD20 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD21", OracleType.Double, 22).Value = (object)this.WD21 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD22", OracleType.Double, 22).Value = (object)this.WD22 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD23", OracleType.Double, 22).Value = (object)this.WD23 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD24", OracleType.Double, 22).Value = (object)this.WD24 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD25", OracleType.Double, 22).Value = (object)this.WD25 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD26", OracleType.Double, 22).Value = (object)this.WD26 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD27", OracleType.Double, 22).Value = (object)this.WD27 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD28", OracleType.Double, 22).Value = (object)this.WD28 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD29", OracleType.Double, 22).Value = (object)this.WD29 ?? DBNull.Value;
                insertCmd.Parameters.Add(":WD30", OracleType.Double, 22).Value = (object)this.WD30 ?? DBNull.Value;
                insertCmd.Parameters.Add(":BZ", OracleType.VarChar, 200).Value = this.BZ;

                insertCmd.ExecuteOracleNonQuery(out RowId);
            }



        }
        public virtual void DbDelete(OracleTransaction trans)
        {
            //DelSql,DelPara
            if (DataState == DataStateType.Delete)
            {
                OracleConnection con = trans.Connection;
                OracleCommand deleteCmd = new OracleCommand();
                deleteCmd.Connection = con;
                deleteCmd.Transaction = trans;
                deleteCmd.CommandText = "DELETE FROM RBWENDU WHERE ROWID='" + RowId.Value + "'";

                deleteCmd.ExecuteNonQuery();
            }

        }
        public virtual void DbUpdate(OracleTransaction trans)
        {
            //UpdateSql,UpdatePara
            if (DataState == DataStateType.Update)
            {
                OracleConnection con = trans.Connection;
                OracleCommand updateCmd = new OracleCommand();
                updateCmd.Connection = con;
                updateCmd.Transaction = trans;
                updateCmd.CommandText = "UPDATE RBWENDU SET GAOLU=:GAOLU,SJ=:SJ,WD1=:WD1,WD2=:WD2,WD3=:WD3,WD4=:WD4,WD5=:WD5,WD6=:WD6,WD7=:WD7,WD8=:WD8,WD9=:WD9,WD10=:WD10,WD11=:WD11,WD12=:WD12,WD13=:WD13,WD14=:WD14,WD15=:WD15,WD16=:WD16,WD17=:WD17,WD18=:WD18,WD19=:WD19,WD20=:WD20,WD21=:WD21,WD22=:WD22,WD23=:WD23,WD24=:WD24,WD25=:WD25,WD26=:WD26,WD27=:WD27,WD28=:WD28,WD29=:WD29,WD30=:WD30,BZ=:BZ WHERE ROWID='" + RowId.Value + "'";
                updateCmd.Parameters.Add(":GAOLU", OracleType.Int32, 22).Value = this.GAOLU;
                updateCmd.Parameters.Add(":SJ", OracleType.DateTime, 7).Value = this.SJ;
                updateCmd.Parameters.Add(":WD1", OracleType.Double, 22).Value = (object)this.WD1 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD2", OracleType.Double, 22).Value = (object)this.WD2 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD3", OracleType.Double, 22).Value = (object)this.WD3 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD4", OracleType.Double, 22).Value = (object)this.WD4 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD5", OracleType.Double, 22).Value = (object)this.WD5 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD6", OracleType.Double, 22).Value = (object)this.WD6 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD7", OracleType.Double, 22).Value = (object)this.WD7 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD8", OracleType.Double, 22).Value = (object)this.WD8 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD9", OracleType.Double, 22).Value = (object)this.WD9 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD10", OracleType.Double, 22).Value = (object)this.WD10 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD11", OracleType.Double, 22).Value = (object)this.WD11 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD12", OracleType.Double, 22).Value = (object)this.WD12 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD13", OracleType.Double, 22).Value = (object)this.WD13 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD14", OracleType.Double, 22).Value = (object)this.WD14 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD15", OracleType.Double, 22).Value = (object)this.WD15 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD16", OracleType.Double, 22).Value = (object)this.WD16 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD17", OracleType.Double, 22).Value = (object)this.WD17 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD18", OracleType.Double, 22).Value = (object)this.WD18 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD19", OracleType.Double, 22).Value = (object)this.WD19 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD20", OracleType.Double, 22).Value = (object)this.WD20 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD21", OracleType.Double, 22).Value = (object)this.WD21 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD22", OracleType.Double, 22).Value = (object)this.WD22 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD23", OracleType.Double, 22).Value = (object)this.WD23 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD24", OracleType.Double, 22).Value = (object)this.WD24 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD25", OracleType.Double, 22).Value = (object)this.WD25 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD26", OracleType.Double, 22).Value = (object)this.WD26 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD27", OracleType.Double, 22).Value = (object)this.WD27 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD28", OracleType.Double, 22).Value = (object)this.WD28 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD29", OracleType.Double, 22).Value = (object)this.WD29 ?? DBNull.Value;
                updateCmd.Parameters.Add(":WD30", OracleType.Double, 22).Value = (object)this.WD30 ?? DBNull.Value;
                updateCmd.Parameters.Add(":BZ", OracleType.VarChar, 200).Value = this.BZ;

                updateCmd.ExecuteNonQuery();
            }


        }
    }
    public partial class RbwenduTable : DbEntityTable<Rbwendu>, IFilter
    {
        private DateTime? loadRq = null;
        private int? loadGaolu = null;
        //Rbwendu,ltznConnectionString
        public void LoadByRQGAOLU(DateTime rq,int gaolu)
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
            selectCmd.CommandText = "SELECT GAOLU,SJ,WD1,WD2,WD3,WD4,WD5,WD6,WD7,WD8,WD9,WD10,WD11,WD12,WD13,WD14,WD15,WD16,WD17,WD18,WD19,WD20,WD21,WD22,WD23,WD24,WD25,WD26,WD27,WD28,WD29,WD30,BZ,ROWID FROM RBWENDU WHERE trunc(SJ)=:RQ and gaolu=:GAOLU";
            selectCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = rq;
            selectCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = gaolu;

            OracleDataReader dr = selectCmd.ExecuteReader();
            while (dr.Read())
            {
                Rbwendu item = new Rbwendu();
                item.BeginInit();
                if (dr.IsDBNull(0)) item.GAOLU = 0; else item.GAOLU = dr.GetInt32(0);
                if (dr.IsDBNull(1)) item.SJ = DateTime.Now; else item.SJ = dr.GetDateTime(1);
                if (dr.IsDBNull(2)) item.WD1 = null; else item.WD1 = dr.GetDouble(2);
                if (dr.IsDBNull(3)) item.WD2 = null; else item.WD2 = dr.GetDouble(3);
                if (dr.IsDBNull(4)) item.WD3 = null; else item.WD3 = dr.GetDouble(4);
                if (dr.IsDBNull(5)) item.WD4 = null; else item.WD4 = dr.GetDouble(5);
                if (dr.IsDBNull(6)) item.WD5 = null; else item.WD5 = dr.GetDouble(6);
                if (dr.IsDBNull(7)) item.WD6 = null; else item.WD6 = dr.GetDouble(7);
                if (dr.IsDBNull(8)) item.WD7 = null; else item.WD7 = dr.GetDouble(8);
                if (dr.IsDBNull(9)) item.WD8 = null; else item.WD8 = dr.GetDouble(9);
                if (dr.IsDBNull(10)) item.WD9 = null; else item.WD9 = dr.GetDouble(10);
                if (dr.IsDBNull(11)) item.WD10 = null; else item.WD10 = dr.GetDouble(11);
                if (dr.IsDBNull(12)) item.WD11 = null; else item.WD11 = dr.GetDouble(12);
                if (dr.IsDBNull(13)) item.WD12 = null; else item.WD12 = dr.GetDouble(13);
                if (dr.IsDBNull(14)) item.WD13 = null; else item.WD13 = dr.GetDouble(14);
                if (dr.IsDBNull(15)) item.WD14 = null; else item.WD14 = dr.GetDouble(15);
                if (dr.IsDBNull(16)) item.WD15 = null; else item.WD15 = dr.GetDouble(16);
                if (dr.IsDBNull(17)) item.WD16 = null; else item.WD16 = dr.GetDouble(17);
                if (dr.IsDBNull(18)) item.WD17 = null; else item.WD17 = dr.GetDouble(18);
                if (dr.IsDBNull(19)) item.WD18 = null; else item.WD18 = dr.GetDouble(19);
                if (dr.IsDBNull(20)) item.WD19 = null; else item.WD19 = dr.GetDouble(20);
                if (dr.IsDBNull(21)) item.WD20 = null; else item.WD20 = dr.GetDouble(21);
                if (dr.IsDBNull(22)) item.WD21 = null; else item.WD21 = dr.GetDouble(22);
                if (dr.IsDBNull(23)) item.WD22 = null; else item.WD22 = dr.GetDouble(23);
                if (dr.IsDBNull(24)) item.WD23 = null; else item.WD23 = dr.GetDouble(24);
                if (dr.IsDBNull(25)) item.WD24 = null; else item.WD24 = dr.GetDouble(25);
                if (dr.IsDBNull(26)) item.WD25 = null; else item.WD25 = dr.GetDouble(26);
                if (dr.IsDBNull(27)) item.WD26 = null; else item.WD26 = dr.GetDouble(27);
                if (dr.IsDBNull(28)) item.WD27 = null; else item.WD27 = dr.GetDouble(28);
                if (dr.IsDBNull(29)) item.WD28 = null; else item.WD28 = dr.GetDouble(29);
                if (dr.IsDBNull(30)) item.WD29 = null; else item.WD29 = dr.GetDouble(30);
                if (dr.IsDBNull(31)) item.WD30 = null; else item.WD30 = dr.GetDouble(31);
                if (dr.IsDBNull(32)) item.BZ = ""; else item.BZ = dr.GetString(32);
                item.RowId = dr.GetString(33);

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
                    item.SJ = new DateTime(loadRq.Value.Year, loadRq.Value.Month, loadRq.Value.Day, item.SJ.Hour, item.SJ.Minute, 0);
                    item.GAOLU = loadGaolu.Value;
                }
            }
            base.Save();
        }
    
    }

}
