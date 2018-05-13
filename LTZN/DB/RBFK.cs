using System;
using System.ComponentModel;
using System.Data.OracleClient;

namespace ZHCDB
{
    public partial class Rbfk : DbEntity
    {
        private System.Nullable<int> _GAOLU;
        private string _FK1 = "";
        private string _FK2 = "";
        private string _FK3 = "";
        private string _FK4 = "";
        private string _FK5 = "";
        private string _FK6 = "";
        private string _FK7 = "";
        private string _FK8 = "";
        private string _FK9 = "";
        private string _FK10 = "";
        private string _FK11 = "";
        private string _FK12 = "";
        private string _FK13 = "";
        private string _FK14 = "";
        private string _YX = "";
        private string _GCY = "";
        private string _BZ = "";
        private System.Nullable<System.DateTime> _RQ = null;
        private System.Nullable<System.DateTime> _SJ = null;
        private string _FK15 = "";
        private string _FK16 = "";
        private string _FK17 = "";
        private string _FK18 = "";
        private string _FK19 = "";
        private string _FK20 = "";
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
        public string FK1
        {
            get
            {
                return this._FK1;
            }
            set
            {
                if (!_FK1.Equals(value))
                {
                    _FK1 = value;
                    RaisePropertyChanged("FK1", true);
                }

            }
        }
        public string FK2
        {
            get
            {
                return this._FK2;
            }
            set
            {
                if (!_FK2.Equals(value))
                {
                    _FK2 = value;
                    RaisePropertyChanged("FK2", true);
                }

            }
        }
        public string FK3
        {
            get
            {
                return this._FK3;
            }
            set
            {
                if (!_FK3.Equals(value))
                {
                    _FK3 = value;
                    RaisePropertyChanged("FK3", true);
                }

            }
        }
        public string FK4
        {
            get
            {
                return this._FK4;
            }
            set
            {
                if (!_FK4.Equals(value))
                {
                    _FK4 = value;
                    RaisePropertyChanged("FK4", true);
                }

            }
        }
        public string FK5
        {
            get
            {
                return this._FK5;
            }
            set
            {
                if (!_FK5.Equals(value))
                {
                    _FK5 = value;
                    RaisePropertyChanged("FK5", true);
                }

            }
        }
        public string FK6
        {
            get
            {
                return this._FK6;
            }
            set
            {
                if (!_FK6.Equals(value))
                {
                    _FK6 = value;
                    RaisePropertyChanged("FK6", true);
                }

            }
        }
        public string FK7
        {
            get
            {
                return this._FK7;
            }
            set
            {
                if (!_FK7.Equals(value))
                {
                    _FK7 = value;
                    RaisePropertyChanged("FK7", true);
                }

            }
        }
        public string FK8
        {
            get
            {
                return this._FK8;
            }
            set
            {
                if (!_FK8.Equals(value))
                {
                    _FK8 = value;
                    RaisePropertyChanged("FK8", true);
                }

            }
        }
        public string FK9
        {
            get
            {
                return this._FK9;
            }
            set
            {
                if (!_FK9.Equals(value))
                {
                    _FK9 = value;
                    RaisePropertyChanged("FK9", true);
                }

            }
        }
        public string FK10
        {
            get
            {
                return this._FK10;
            }
            set
            {
                if (!_FK10.Equals(value))
                {
                    _FK10 = value;
                    RaisePropertyChanged("FK10", true);
                }

            }
        }
        public string FK11
        {
            get
            {
                return this._FK11;
            }
            set
            {
                if (!_FK11.Equals(value))
                {
                    _FK11 = value;
                    RaisePropertyChanged("FK11", true);
                }

            }
        }
        public string FK12
        {
            get
            {
                return this._FK12;
            }
            set
            {
                if (!_FK12.Equals(value))
                {
                    _FK12 = value;
                    RaisePropertyChanged("FK12", true);
                }

            }
        }
        public string FK13
        {
            get
            {
                return this._FK13;
            }
            set
            {
                if (!_FK13.Equals(value))
                {
                    _FK13 = value;
                    RaisePropertyChanged("FK13", true);
                }

            }
        }
        public string FK14
        {
            get
            {
                return this._FK14;
            }
            set
            {
                if (!_FK14.Equals(value))
                {
                    _FK14 = value;
                    RaisePropertyChanged("FK14", true);
                }

            }
        }
        public string YX
        {
            get
            {
                return this._YX;
            }
            set
            {
                if (!_YX.Equals(value))
                {
                    _YX = value;
                    RaisePropertyChanged("YX", true);
                }

            }
        }
        public string GCY
        {
            get
            {
                return this._GCY;
            }
            set
            {
                if (!_GCY.Equals(value))
                {
                    _GCY = value;
                    RaisePropertyChanged("GCY", true);
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
        public string FK15
        {
            get
            {
                return this._FK15;
            }
            set
            {
                if (!_FK15.Equals(value))
                {
                    _FK15 = value;
                    RaisePropertyChanged("FK15", true);
                }

            }
        }
        public string FK16
        {
            get
            {
                return this._FK16;
            }
            set
            {
                if (!_FK16.Equals(value))
                {
                    _FK16 = value;
                    RaisePropertyChanged("FK16", true);
                }

            }
        }
        public string FK17
        {
            get
            {
                return this._FK17;
            }
            set
            {
                if (!_FK17.Equals(value))
                {
                    _FK17 = value;
                    RaisePropertyChanged("FK17", true);
                }

            }
        }
        public string FK18
        {
            get
            {
                return this._FK18;
            }
            set
            {
                if (!_FK18.Equals(value))
                {
                    _FK18 = value;
                    RaisePropertyChanged("FK18", true);
                }

            }
        }
        public string FK19
        {
            get
            {
                return this._FK19;
            }
            set
            {
                if (!_FK19.Equals(value))
                {
                    _FK19 = value;
                    RaisePropertyChanged("FK19", true);
                }

            }
        }
        public string FK20
        {
            get
            {
                return this._FK20;
            }
            set
            {
                if (!_FK20.Equals(value))
                {
                    _FK20 = value;
                    RaisePropertyChanged("FK20", true);
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
                insertCmd.CommandText = "INSERT INTO RBFK(GAOLU,FK1,FK2,FK3,FK4,FK5,FK6,FK7,FK8,FK9,FK10,FK11,FK12,FK13,FK14,YX,GCY,BZ,RQ,SJ,FK15,FK16,FK17,FK18,FK19,FK20) VALUES(:GAOLU,:FK1,:FK2,:FK3,:FK4,:FK5,:FK6,:FK7,:FK8,:FK9,:FK10,:FK11,:FK12,:FK13,:FK14,:YX,:GCY,:BZ,:RQ,:SJ,:FK15,:FK16,:FK17,:FK18,:FK19,:FK20)";
                insertCmd.Parameters.Add(":GAOLU", OracleType.Byte, 22).Value = (object)this.GAOLU ?? DBNull.Value;
                insertCmd.Parameters.Add(":FK1", OracleType.VarChar, 4).Value = this.FK1;
                insertCmd.Parameters.Add(":FK2", OracleType.VarChar, 4).Value = this.FK2;
                insertCmd.Parameters.Add(":FK3", OracleType.VarChar, 4).Value = this.FK3;
                insertCmd.Parameters.Add(":FK4", OracleType.VarChar, 4).Value = this.FK4;
                insertCmd.Parameters.Add(":FK5", OracleType.VarChar, 4).Value = this.FK5;
                insertCmd.Parameters.Add(":FK6", OracleType.VarChar, 4).Value = this.FK6;
                insertCmd.Parameters.Add(":FK7", OracleType.VarChar, 4).Value = this.FK7;
                insertCmd.Parameters.Add(":FK8", OracleType.VarChar, 4).Value = this.FK8;
                insertCmd.Parameters.Add(":FK9", OracleType.VarChar, 4).Value = this.FK9;
                insertCmd.Parameters.Add(":FK10", OracleType.VarChar, 4).Value = this.FK10;
                insertCmd.Parameters.Add(":FK11", OracleType.VarChar, 4).Value = this.FK11;
                insertCmd.Parameters.Add(":FK12", OracleType.VarChar, 4).Value = this.FK12;
                insertCmd.Parameters.Add(":FK13", OracleType.VarChar, 4).Value = this.FK13;
                insertCmd.Parameters.Add(":FK14", OracleType.VarChar, 4).Value = this.FK14;
                insertCmd.Parameters.Add(":YX", OracleType.VarChar, 10).Value = this.YX;
                insertCmd.Parameters.Add(":GCY", OracleType.VarChar, 10).Value = this.GCY;
                insertCmd.Parameters.Add(":BZ", OracleType.VarChar, 100).Value = this.BZ;
                insertCmd.Parameters.Add(":RQ", OracleType.DateTime, 7).Value = (object)this.RQ ?? DBNull.Value;
                insertCmd.Parameters.Add(":SJ", OracleType.DateTime, 7).Value = (object)this.SJ ?? DBNull.Value;
                insertCmd.Parameters.Add(":FK15", OracleType.VarChar, 4).Value = this.FK15;
                insertCmd.Parameters.Add(":FK16", OracleType.VarChar, 4).Value = this.FK16;
                insertCmd.Parameters.Add(":FK17", OracleType.VarChar, 4).Value = this.FK17;
                insertCmd.Parameters.Add(":FK18", OracleType.VarChar, 4).Value = this.FK18;
                insertCmd.Parameters.Add(":FK19", OracleType.VarChar, 4).Value = this.FK19;
                insertCmd.Parameters.Add(":FK20", OracleType.VarChar, 4).Value = this.FK20;

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
                deleteCmd.CommandText = "DELETE FROM RBFK WHERE ROWID='" + RowId.Value + "'";

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
                updateCmd.CommandText = "UPDATE RBFK SET GAOLU=:GAOLU,FK1=:FK1,FK2=:FK2,FK3=:FK3,FK4=:FK4,FK5=:FK5,FK6=:FK6,FK7=:FK7,FK8=:FK8,FK9=:FK9,FK10=:FK10,FK11=:FK11,FK12=:FK12,FK13=:FK13,FK14=:FK14,YX=:YX,GCY=:GCY,BZ=:BZ,RQ=:RQ,SJ=:SJ,FK15=:FK15,FK16=:FK16,FK17=:FK17,FK18=:FK18,FK19=:FK19,FK20=:FK20 WHERE ROWID='" + RowId.Value + "'";
                updateCmd.Parameters.Add(":GAOLU", OracleType.Byte, 22).Value = (object)this.GAOLU ?? DBNull.Value;
                updateCmd.Parameters.Add(":FK1", OracleType.VarChar, 4).Value = this.FK1;
                updateCmd.Parameters.Add(":FK2", OracleType.VarChar, 4).Value = this.FK2;
                updateCmd.Parameters.Add(":FK3", OracleType.VarChar, 4).Value = this.FK3;
                updateCmd.Parameters.Add(":FK4", OracleType.VarChar, 4).Value = this.FK4;
                updateCmd.Parameters.Add(":FK5", OracleType.VarChar, 4).Value = this.FK5;
                updateCmd.Parameters.Add(":FK6", OracleType.VarChar, 4).Value = this.FK6;
                updateCmd.Parameters.Add(":FK7", OracleType.VarChar, 4).Value = this.FK7;
                updateCmd.Parameters.Add(":FK8", OracleType.VarChar, 4).Value = this.FK8;
                updateCmd.Parameters.Add(":FK9", OracleType.VarChar, 4).Value = this.FK9;
                updateCmd.Parameters.Add(":FK10", OracleType.VarChar, 4).Value = this.FK10;
                updateCmd.Parameters.Add(":FK11", OracleType.VarChar, 4).Value = this.FK11;
                updateCmd.Parameters.Add(":FK12", OracleType.VarChar, 4).Value = this.FK12;
                updateCmd.Parameters.Add(":FK13", OracleType.VarChar, 4).Value = this.FK13;
                updateCmd.Parameters.Add(":FK14", OracleType.VarChar, 4).Value = this.FK14;
                updateCmd.Parameters.Add(":YX", OracleType.VarChar, 10).Value = this.YX;
                updateCmd.Parameters.Add(":GCY", OracleType.VarChar, 10).Value = this.GCY;
                updateCmd.Parameters.Add(":BZ", OracleType.VarChar, 100).Value = this.BZ;
                updateCmd.Parameters.Add(":RQ", OracleType.DateTime, 7).Value = (object)this.RQ ?? DBNull.Value;
                updateCmd.Parameters.Add(":SJ", OracleType.DateTime, 7).Value = (object)this.SJ ?? DBNull.Value;
                updateCmd.Parameters.Add(":FK15", OracleType.VarChar, 4).Value = this.FK15;
                updateCmd.Parameters.Add(":FK16", OracleType.VarChar, 4).Value = this.FK16;
                updateCmd.Parameters.Add(":FK17", OracleType.VarChar, 4).Value = this.FK17;
                updateCmd.Parameters.Add(":FK18", OracleType.VarChar, 4).Value = this.FK18;
                updateCmd.Parameters.Add(":FK19", OracleType.VarChar, 4).Value = this.FK19;
                updateCmd.Parameters.Add(":FK20", OracleType.VarChar, 4).Value = this.FK20;

                updateCmd.ExecuteNonQuery();
            }


        }
    }
    public partial class RbfkTable : DbEntityTable<Rbfk>, IFilter
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
            selectCmd.CommandText = "SELECT GAOLU,FK1,FK2,FK3,FK4,FK5,FK6,FK7,FK8,FK9,FK10,FK11,FK12,FK13,FK14,YX,GCY,BZ,RQ,SJ,FK15,FK16,FK17,FK18,FK19,FK20,ROWID FROM RBFK WHERE   (TRUNC(SJ) = :RQ) AND (GAOLU = :GAOLU) order by SJ";
            selectCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = rq;
            selectCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = gaolu;

            OracleDataReader dr = selectCmd.ExecuteReader();
            while (dr.Read())
            {
                Rbfk item = new Rbfk();
                item.BeginInit();
                if (dr.IsDBNull(0)) item.GAOLU = null; else item.GAOLU = dr.GetInt32(0);
                if (dr.IsDBNull(1)) item.FK1 = ""; else item.FK1 = dr.GetString(1);
                if (dr.IsDBNull(2)) item.FK2 = ""; else item.FK2 = dr.GetString(2);
                if (dr.IsDBNull(3)) item.FK3 = ""; else item.FK3 = dr.GetString(3);
                if (dr.IsDBNull(4)) item.FK4 = ""; else item.FK4 = dr.GetString(4);
                if (dr.IsDBNull(5)) item.FK5 = ""; else item.FK5 = dr.GetString(5);
                if (dr.IsDBNull(6)) item.FK6 = ""; else item.FK6 = dr.GetString(6);
                if (dr.IsDBNull(7)) item.FK7 = ""; else item.FK7 = dr.GetString(7);
                if (dr.IsDBNull(8)) item.FK8 = ""; else item.FK8 = dr.GetString(8);
                if (dr.IsDBNull(9)) item.FK9 = ""; else item.FK9 = dr.GetString(9);
                if (dr.IsDBNull(10)) item.FK10 = ""; else item.FK10 = dr.GetString(10);
                if (dr.IsDBNull(11)) item.FK11 = ""; else item.FK11 = dr.GetString(11);
                if (dr.IsDBNull(12)) item.FK12 = ""; else item.FK12 = dr.GetString(12);
                if (dr.IsDBNull(13)) item.FK13 = ""; else item.FK13 = dr.GetString(13);
                if (dr.IsDBNull(14)) item.FK14 = ""; else item.FK14 = dr.GetString(14);
                if (dr.IsDBNull(15)) item.YX = ""; else item.YX = dr.GetString(15);
                if (dr.IsDBNull(16)) item.GCY = ""; else item.GCY = dr.GetString(16);
                if (dr.IsDBNull(17)) item.BZ = ""; else item.BZ = dr.GetString(17);
                if (dr.IsDBNull(18)) item.RQ = null; else item.RQ = dr.GetDateTime(18);
                if (dr.IsDBNull(19)) item.SJ = null; else item.SJ = dr.GetDateTime(19);
                if (dr.IsDBNull(20)) item.FK15 = ""; else item.FK15 = dr.GetString(20);
                if (dr.IsDBNull(21)) item.FK16 = ""; else item.FK16 = dr.GetString(21);
                if (dr.IsDBNull(22)) item.FK17 = ""; else item.FK17 = dr.GetString(22);
                if (dr.IsDBNull(23)) item.FK18 = ""; else item.FK18 = dr.GetString(23);
                if (dr.IsDBNull(24)) item.FK19 = ""; else item.FK19 = dr.GetString(24);
                if (dr.IsDBNull(25)) item.FK20 = ""; else item.FK20 = dr.GetString(25);
                item.RowId = dr.GetString(26);

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
                        item.RQ = new DateTime(loadRq.Value.Year, loadRq.Value.Month, loadRq.Value.Day, DateTime.Now.Hour, DateTime.Now.Minute, 0);

                    item.RQ = new DateTime(loadRq.Value.Year, loadRq.Value.Month, loadRq.Value.Day, item.RQ.Value.Hour, item.RQ.Value.Minute, 0);

                    item.GAOLU = loadGaolu.Value;
                    item.SJ = item.RQ;
                }
            }
            base.Save();
        }
    }

  
}
