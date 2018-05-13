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
    public partial class Rblq : DbEntity
    {
        private System.Nullable<int> _GAOLU;
        private string _BC = "";
        private string _CC = "";
        private System.Nullable<double> _HAO1;
        private System.Nullable<System.DateTime> _HAO2 = null;
        private System.Nullable<double> _HAO3;
        private System.Nullable<double> _HAO4;
        private System.Nullable<double> _HAO5;
        private System.Nullable<double> _HAO6;
        private System.Nullable<double> _HAO7;
        private System.Nullable<double> _HAO8;
        private System.Nullable<double> _HAO9;
        private System.Nullable<double> _HAO10;
        private System.Nullable<double> _HAO11;
        private System.Nullable<double> _HAO12;
        private System.Nullable<double> _HAO13;
        private System.Nullable<double> _HAO14;
        private System.Nullable<double> _HAO15;
        private System.Nullable<double> _HAO16;
        private System.Nullable<double> _HAO17;
        private System.Nullable<double> _HAO18;
        private System.Nullable<double> _HAO19;
        private System.Nullable<double> _HAO20;
        private System.Nullable<double> _HAO21;
        private System.Nullable<double> _HAO22;
        private System.Nullable<double> _HAO23;
        private System.Nullable<double> _HAO24;
        private System.Nullable<double> _HAO25;
        private System.Nullable<double> _HAO26;
        private System.Nullable<double> _HAO27;
        private System.Nullable<double> _HAO28;
        private string _BZ = "";
        private System.Nullable<System.DateTime> _SJ = null;
        private System.Nullable<double> _PJZ;
        private System.Nullable<double> _HAO29;
        private System.Nullable<double> _HAO30;
        private System.Nullable<double> _HAO31;
        private System.Nullable<double> _HAO32;
        private System.Nullable<double> _HAO33;
        private System.Nullable<double> _HAO34;
        private System.Nullable<double> _HAO35;
        private System.Nullable<double> _HAO36;
        private System.Nullable<double> _HAO37;
        private System.Nullable<double> _HAO38;
        private System.Nullable<double> _HAO39;
        private System.Nullable<double> _HAO40;
        private System.Nullable<double> _HAO41;
        private System.Nullable<double> _HAO42;
        private System.Nullable<double> _HAO43;
        private System.Nullable<double> _HAO44;
        private System.Nullable<double> _HAO45;
        private System.Nullable<double> _HAO46;
        private System.Nullable<double> _HAO47;
        private System.Nullable<double> _HAO48;
        private System.Nullable<double> _HAO49;
        private System.Nullable<double> _HAO50;
        private System.Nullable<double> _CYS;
        private System.Nullable<double> _GYS;
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
        public string CC
        {
            get
            {
                return this._CC;
            }
            set
            {
                if (!_CC.Equals(value))
                {
                    _CC = value;
                    RaisePropertyChanged("CC", true);
                }

            }
        }
        /// <summary>
        /// 常压水水压kPa
        /// </summary>
        public System.Nullable<double> CYS
        {
            get
            {
                return this._CYS;
            }
            set
            {
                if (!_CYS.Equals(value))
                {
                    _CYS = value;
                    RaisePropertyChanged("CYS", true);
                }

            }
        }
        /// <summary>
        /// 高压水水压kPa
        /// </summary>
        public System.Nullable<double> GYS
        {
            get
            {
                return this._GYS;
            }
            set
            {
                if (!_GYS.Equals(value))
                {
                    _GYS = value;
                    RaisePropertyChanged("GYS", true);
                }

            }
        }

        public System.Nullable<double> HAO1
        {
            get
            {
                return this._HAO1;
            }
            set
            {
                if (!_HAO1.Equals(value))
                {
                    _HAO1 = value;
                    RaisePropertyChanged("HAO1", true);
                }

            }
        }
        public System.Nullable<System.DateTime> HAO2
        {
            get
            {
                return this._HAO2;
            }
            set
            {
                if (!_HAO2.Equals(value))
                {
                    _HAO2 = value;
                    RaisePropertyChanged("HAO2", true);
                }

            }
        }
        public System.Nullable<double> HAO3
        {
            get
            {
                return this._HAO3;
            }
            set
            {
                if (!_HAO3.Equals(value))
                {
                    _HAO3 = value;
                    RaisePropertyChanged("HAO3", true);
                }

            }
        }
        public System.Nullable<double> HAO4
        {
            get
            {
                return this._HAO4;
            }
            set
            {
                if (!_HAO4.Equals(value))
                {
                    _HAO4 = value;
                    RaisePropertyChanged("HAO4", true);
                }

            }
        }
        public System.Nullable<double> HAO5
        {
            get
            {
                return this._HAO5;
            }
            set
            {
                if (!_HAO5.Equals(value))
                {
                    _HAO5 = value;
                    RaisePropertyChanged("HAO5", true);
                }

            }
        }
        public System.Nullable<double> HAO6
        {
            get
            {
                return this._HAO6;
            }
            set
            {
                if (!_HAO6.Equals(value))
                {
                    _HAO6 = value;
                    RaisePropertyChanged("HAO6", true);
                }

            }
        }
        public System.Nullable<double> HAO7
        {
            get
            {
                return this._HAO7;
            }
            set
            {
                if (!_HAO7.Equals(value))
                {
                    _HAO7 = value;
                    RaisePropertyChanged("HAO7", true);
                }

            }
        }
        public System.Nullable<double> HAO8
        {
            get
            {
                return this._HAO8;
            }
            set
            {
                if (!_HAO8.Equals(value))
                {
                    _HAO8 = value;
                    RaisePropertyChanged("HAO8", true);
                }

            }
        }
        public System.Nullable<double> HAO9
        {
            get
            {
                return this._HAO9;
            }
            set
            {
                if (!_HAO9.Equals(value))
                {
                    _HAO9 = value;
                    RaisePropertyChanged("HAO9", true);
                }

            }
        }
        public System.Nullable<double> HAO10
        {
            get
            {
                return this._HAO10;
            }
            set
            {
                if (!_HAO10.Equals(value))
                {
                    _HAO10 = value;
                    RaisePropertyChanged("HAO10", true);
                }

            }
        }
        public System.Nullable<double> HAO11
        {
            get
            {
                return this._HAO11;
            }
            set
            {
                if (!_HAO11.Equals(value))
                {
                    _HAO11 = value;
                    RaisePropertyChanged("HAO11", true);
                }

            }
        }
        public System.Nullable<double> HAO12
        {
            get
            {
                return this._HAO12;
            }
            set
            {
                if (!_HAO12.Equals(value))
                {
                    _HAO12 = value;
                    RaisePropertyChanged("HAO12", true);
                }

            }
        }
        public System.Nullable<double> HAO13
        {
            get
            {
                return this._HAO13;
            }
            set
            {
                if (!_HAO13.Equals(value))
                {
                    _HAO13 = value;
                    RaisePropertyChanged("HAO13", true);
                }

            }
        }
        public System.Nullable<double> HAO14
        {
            get
            {
                return this._HAO14;
            }
            set
            {
                if (!_HAO14.Equals(value))
                {
                    _HAO14 = value;
                    RaisePropertyChanged("HAO14", true);
                }

            }
        }
        public System.Nullable<double> HAO15
        {
            get
            {
                return this._HAO15;
            }
            set
            {
                if (!_HAO15.Equals(value))
                {
                    _HAO15 = value;
                    RaisePropertyChanged("HAO15", true);
                }

            }
        }
        public System.Nullable<double> HAO16
        {
            get
            {
                return this._HAO16;
            }
            set
            {
                if (!_HAO16.Equals(value))
                {
                    _HAO16 = value;
                    RaisePropertyChanged("HAO16", true);
                }

            }
        }
        public System.Nullable<double> HAO17
        {
            get
            {
                return this._HAO17;
            }
            set
            {
                if (!_HAO17.Equals(value))
                {
                    _HAO17 = value;
                    RaisePropertyChanged("HAO17", true);
                }

            }
        }
        public System.Nullable<double> HAO18
        {
            get
            {
                return this._HAO18;
            }
            set
            {
                if (!_HAO18.Equals(value))
                {
                    _HAO18 = value;
                    RaisePropertyChanged("HAO18", true);
                }

            }
        }
        public System.Nullable<double> HAO19
        {
            get
            {
                return this._HAO19;
            }
            set
            {
                if (!_HAO19.Equals(value))
                {
                    _HAO19 = value;
                    RaisePropertyChanged("HAO19", true);
                }

            }
        }
        public System.Nullable<double> HAO20
        {
            get
            {
                return this._HAO20;
            }
            set
            {
                if (!_HAO20.Equals(value))
                {
                    _HAO20 = value;
                    RaisePropertyChanged("HAO20", true);
                }

            }
        }
        public System.Nullable<double> HAO21
        {
            get
            {
                return this._HAO21;
            }
            set
            {
                if (!_HAO21.Equals(value))
                {
                    _HAO21 = value;
                    RaisePropertyChanged("HAO21", true);
                }

            }
        }
        public System.Nullable<double> HAO22
        {
            get
            {
                return this._HAO22;
            }
            set
            {
                if (!_HAO22.Equals(value))
                {
                    _HAO22 = value;
                    RaisePropertyChanged("HAO22", true);
                }

            }
        }
        public System.Nullable<double> HAO23
        {
            get
            {
                return this._HAO23;
            }
            set
            {
                if (!_HAO23.Equals(value))
                {
                    _HAO23 = value;
                    RaisePropertyChanged("HAO23", true);
                }

            }
        }
        public System.Nullable<double> HAO24
        {
            get
            {
                return this._HAO24;
            }
            set
            {
                if (!_HAO24.Equals(value))
                {
                    _HAO24 = value;
                    RaisePropertyChanged("HAO24", true);
                }

            }
        }
        public System.Nullable<double> HAO25
        {
            get
            {
                return this._HAO25;
            }
            set
            {
                if (!_HAO25.Equals(value))
                {
                    _HAO25 = value;
                    RaisePropertyChanged("HAO25", true);
                }

            }
        }
        public System.Nullable<double> HAO26
        {
            get
            {
                return this._HAO26;
            }
            set
            {
                if (!_HAO26.Equals(value))
                {
                    _HAO26 = value;
                    RaisePropertyChanged("HAO26", true);
                }

            }
        }
        public System.Nullable<double> HAO27
        {
            get
            {
                return this._HAO27;
            }
            set
            {
                if (!_HAO27.Equals(value))
                {
                    _HAO27 = value;
                    RaisePropertyChanged("HAO27", true);
                }

            }
        }
        public System.Nullable<double> HAO28
        {
            get
            {
                return this._HAO28;
            }
            set
            {
                if (!_HAO28.Equals(value))
                {
                    _HAO28 = value;
                    RaisePropertyChanged("HAO28", true);
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
        public System.Nullable<double> PJZ
        {
            get
            {
                return this._PJZ;
            }
            set
            {
                if (!_PJZ.Equals(value))
                {
                    _PJZ = value;
                    RaisePropertyChanged("PJZ", true);
                }

            }
        }
        public System.Nullable<double> HAO29
        {
            get
            {
                return this._HAO29;
            }
            set
            {
                if (!_HAO29.Equals(value))
                {
                    _HAO29 = value;
                    RaisePropertyChanged("HAO29", true);
                }

            }
        }
        public System.Nullable<double> HAO30
        {
            get
            {
                return this._HAO30;
            }
            set
            {
                if (!_HAO30.Equals(value))
                {
                    _HAO30 = value;
                    RaisePropertyChanged("HAO30", true);
                }

            }
        }
        public System.Nullable<double> HAO31
        {
            get
            {
                return this._HAO31;
            }
            set
            {
                if (!_HAO31.Equals(value))
                {
                    _HAO31 = value;
                    RaisePropertyChanged("HAO31", true);
                }

            }
        }
        public System.Nullable<double> HAO32
        {
            get
            {
                return this._HAO32;
            }
            set
            {
                if (!_HAO32.Equals(value))
                {
                    _HAO32 = value;
                    RaisePropertyChanged("HAO32", true);
                }

            }
        }
        public System.Nullable<double> HAO33
        {
            get
            {
                return this._HAO33;
            }
            set
            {
                if (!_HAO33.Equals(value))
                {
                    _HAO33 = value;
                    RaisePropertyChanged("HAO33", true);
                }

            }
        }
        public System.Nullable<double> HAO34
        {
            get
            {
                return this._HAO34;
            }
            set
            {
                if (!_HAO34.Equals(value))
                {
                    _HAO34 = value;
                    RaisePropertyChanged("HAO34", true);
                }

            }
        }
        public System.Nullable<double> HAO35
        {
            get
            {
                return this._HAO35;
            }
            set
            {
                if (!_HAO35.Equals(value))
                {
                    _HAO35 = value;
                    RaisePropertyChanged("HAO35", true);
                }

            }
        }
        public System.Nullable<double> HAO36
        {
            get
            {
                return this._HAO36;
            }
            set
            {
                if (!_HAO36.Equals(value))
                {
                    _HAO36 = value;
                    RaisePropertyChanged("HAO36", true);
                }

            }
        }
        public System.Nullable<double> HAO37
        {
            get
            {
                return this._HAO37;
            }
            set
            {
                if (!_HAO37.Equals(value))
                {
                    _HAO37 = value;
                    RaisePropertyChanged("HAO37", true);
                }

            }
        }
        public System.Nullable<double> HAO38
        {
            get
            {
                return this._HAO38;
            }
            set
            {
                if (!_HAO38.Equals(value))
                {
                    _HAO38 = value;
                    RaisePropertyChanged("HAO38", true);
                }

            }
        }
        public System.Nullable<double> HAO39
        {
            get
            {
                return this._HAO39;
            }
            set
            {
                if (!_HAO39.Equals(value))
                {
                    _HAO39 = value;
                    RaisePropertyChanged("HAO39", true);
                }

            }
        }
        public System.Nullable<double> HAO40
        {
            get
            {
                return this._HAO40;
            }
            set
            {
                if (!_HAO40.Equals(value))
                {
                    _HAO40 = value;
                    RaisePropertyChanged("HAO40", true);
                }

            }
        }
        public System.Nullable<double> HAO41
        {
            get
            {
                return this._HAO41;
            }
            set
            {
                if (!_HAO41.Equals(value))
                {
                    _HAO41 = value;
                    RaisePropertyChanged("HAO41", true);
                }

            }
        }
        public System.Nullable<double> HAO42
        {
            get
            {
                return this._HAO42;
            }
            set
            {
                if (!_HAO42.Equals(value))
                {
                    _HAO42 = value;
                    RaisePropertyChanged("HAO42", true);
                }

            }
        }
        public System.Nullable<double> HAO43
        {
            get
            {
                return this._HAO43;
            }
            set
            {
                if (!_HAO43.Equals(value))
                {
                    _HAO43 = value;
                    RaisePropertyChanged("HAO43", true);
                }

            }
        }
        public System.Nullable<double> HAO44
        {
            get
            {
                return this._HAO44;
            }
            set
            {
                if (!_HAO44.Equals(value))
                {
                    _HAO44 = value;
                    RaisePropertyChanged("HAO44", true);
                }

            }
        }
        public System.Nullable<double> HAO45
        {
            get
            {
                return this._HAO45;
            }
            set
            {
                if (!_HAO45.Equals(value))
                {
                    _HAO45 = value;
                    RaisePropertyChanged("HAO45", true);
                }

            }
        }
        public System.Nullable<double> HAO46
        {
            get
            {
                return this._HAO46;
            }
            set
            {
                if (!_HAO46.Equals(value))
                {
                    _HAO46 = value;
                    RaisePropertyChanged("HAO46", true);
                }

            }
        }
        public System.Nullable<double> HAO47
        {
            get
            {
                return this._HAO47;
            }
            set
            {
                if (!_HAO47.Equals(value))
                {
                    _HAO47 = value;
                    RaisePropertyChanged("HAO47", true);
                }

            }
        }
        public System.Nullable<double> HAO48
        {
            get
            {
                return this._HAO48;
            }
            set
            {
                if (!_HAO48.Equals(value))
                {
                    _HAO48 = value;
                    RaisePropertyChanged("HAO48", true);
                }

            }
        }
        public System.Nullable<double> HAO49
        {
            get
            {
                return this._HAO49;
            }
            set
            {
                if (!_HAO49.Equals(value))
                {
                    _HAO49 = value;
                    RaisePropertyChanged("HAO49", true);
                }

            }
        }
        public System.Nullable<double> HAO50
        {
            get
            {
                return this._HAO50;
            }
            set
            {
                if (!_HAO50.Equals(value))
                {
                    _HAO50 = value;
                    RaisePropertyChanged("HAO50", true);
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
                insertCmd.CommandText = "INSERT INTO RBLQ(GAOLU,BC,CC,HAO1,HAO2,HAO3,HAO4,HAO5,HAO6,HAO7,HAO8,HAO9,HAO10,HAO11,HAO12,HAO13,HAO14,HAO15,HAO16,HAO17,HAO18,HAO19,HAO20,HAO21,HAO22,HAO23,HAO24,HAO25,HAO26,HAO27,HAO28,BZ,SJ,PJZ,HAO29,HAO30,HAO31,HAO32,HAO33,HAO34,HAO35,HAO36,HAO37,HAO38,HAO39,HAO40,HAO41,HAO42,HAO43,HAO44,HAO45,HAO46,HAO47,HAO48,HAO49,HAO50,CYS,GYS) VALUES(:GAOLU,:BC,:CC,:HAO1,:HAO2,:HAO3,:HAO4,:HAO5,:HAO6,:HAO7,:HAO8,:HAO9,:HAO10,:HAO11,:HAO12,:HAO13,:HAO14,:HAO15,:HAO16,:HAO17,:HAO18,:HAO19,:HAO20,:HAO21,:HAO22,:HAO23,:HAO24,:HAO25,:HAO26,:HAO27,:HAO28,:BZ,:SJ,:PJZ,:HAO29,:HAO30,:HAO31,:HAO32,:HAO33,:HAO34,:HAO35,:HAO36,:HAO37,:HAO38,:HAO39,:HAO40,:HAO41,:HAO42,:HAO43,:HAO44,:HAO45,:HAO46,:HAO47,:HAO48,:HAO49,:HAO50,:CYS,:GYS)";
                insertCmd.Parameters.Add(":GAOLU", OracleType.Int32, 22).Value = (object)this.GAOLU ?? DBNull.Value;
                insertCmd.Parameters.Add(":BC", OracleType.VarChar, 10).Value = this.BC;
                insertCmd.Parameters.Add(":CC", OracleType.VarChar, 8).Value = this.CC;
                insertCmd.Parameters.Add(":HAO1", OracleType.Double, 22).Value = (object)this.HAO1 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO2", OracleType.DateTime, 7).Value = (object)this.HAO2 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO3", OracleType.Double, 22).Value = (object)this.HAO3 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO4", OracleType.Double, 22).Value = (object)this.HAO4 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO5", OracleType.Double, 22).Value = (object)this.HAO5 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO6", OracleType.Double, 22).Value = (object)this.HAO6 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO7", OracleType.Double, 22).Value = (object)this.HAO7 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO8", OracleType.Double, 22).Value = (object)this.HAO8 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO9", OracleType.Double, 22).Value = (object)this.HAO9 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO10", OracleType.Double, 22).Value = (object)this.HAO10 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO11", OracleType.Double, 22).Value = (object)this.HAO11 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO12", OracleType.Double, 22).Value = (object)this.HAO12 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO13", OracleType.Double, 22).Value = (object)this.HAO13 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO14", OracleType.Double, 22).Value = (object)this.HAO14 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO15", OracleType.Double, 22).Value = (object)this.HAO15 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO16", OracleType.Double, 22).Value = (object)this.HAO16 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO17", OracleType.Double, 22).Value = (object)this.HAO17 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO18", OracleType.Double, 22).Value = (object)this.HAO18 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO19", OracleType.Double, 22).Value = (object)this.HAO19 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO20", OracleType.Double, 22).Value = (object)this.HAO20 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO21", OracleType.Double, 22).Value = (object)this.HAO21 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO22", OracleType.Double, 22).Value = (object)this.HAO22 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO23", OracleType.Double, 22).Value = (object)this.HAO23 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO24", OracleType.Double, 22).Value = (object)this.HAO24 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO25", OracleType.Double, 22).Value = (object)this.HAO25 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO26", OracleType.Double, 22).Value = (object)this.HAO26 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO27", OracleType.Double, 22).Value = (object)this.HAO27 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO28", OracleType.Double, 22).Value = (object)this.HAO28 ?? DBNull.Value;
                insertCmd.Parameters.Add(":BZ", OracleType.VarChar, 20).Value = this.BZ;
                insertCmd.Parameters.Add(":SJ", OracleType.DateTime, 7).Value = (object)this.SJ ?? DBNull.Value;
                insertCmd.Parameters.Add(":PJZ", OracleType.Double, 22).Value = (object)this.PJZ ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO29", OracleType.Double, 22).Value = (object)this.HAO29 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO30", OracleType.Double, 22).Value = (object)this.HAO30 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO31", OracleType.Double, 22).Value = (object)this.HAO31 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO32", OracleType.Double, 22).Value = (object)this.HAO32 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO33", OracleType.Double, 22).Value = (object)this.HAO33 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO34", OracleType.Double, 22).Value = (object)this.HAO34 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO35", OracleType.Double, 22).Value = (object)this.HAO35 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO36", OracleType.Double, 22).Value = (object)this.HAO36 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO37", OracleType.Double, 22).Value = (object)this.HAO37 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO38", OracleType.Double, 22).Value = (object)this.HAO38 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO39", OracleType.Double, 22).Value = (object)this.HAO39 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO40", OracleType.Double, 22).Value = (object)this.HAO40 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO41", OracleType.Double, 22).Value = (object)this.HAO41 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO42", OracleType.Double, 22).Value = (object)this.HAO42 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO43", OracleType.Double, 22).Value = (object)this.HAO43 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO44", OracleType.Double, 22).Value = (object)this.HAO44 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO45", OracleType.Double, 22).Value = (object)this.HAO45 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO46", OracleType.Double, 22).Value = (object)this.HAO46 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO47", OracleType.Double, 22).Value = (object)this.HAO47 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO48", OracleType.Double, 22).Value = (object)this.HAO48 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO49", OracleType.Double, 22).Value = (object)this.HAO49 ?? DBNull.Value;
                insertCmd.Parameters.Add(":HAO50", OracleType.Double, 22).Value = (object)this.HAO50 ?? DBNull.Value;
                insertCmd.Parameters.Add(":CYS", OracleType.Double, 22).Value = (object)this.CYS ?? DBNull.Value;
                insertCmd.Parameters.Add(":GYS", OracleType.Double, 22).Value = (object)this.GYS ?? DBNull.Value;
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
                deleteCmd.CommandText = "DELETE FROM RBLQ WHERE ROWID='" + RowId.Value + "'";

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
                updateCmd.CommandText = "UPDATE RBLQ SET GAOLU=:GAOLU,BC=:BC,CC=:CC,HAO1=:HAO1,HAO2=:HAO2,HAO3=:HAO3,HAO4=:HAO4,HAO5=:HAO5,HAO6=:HAO6,HAO7=:HAO7,HAO8=:HAO8,HAO9=:HAO9,HAO10=:HAO10,HAO11=:HAO11,HAO12=:HAO12,HAO13=:HAO13,HAO14=:HAO14,HAO15=:HAO15,HAO16=:HAO16,HAO17=:HAO17,HAO18=:HAO18,HAO19=:HAO19,HAO20=:HAO20,HAO21=:HAO21,HAO22=:HAO22,HAO23=:HAO23,HAO24=:HAO24,HAO25=:HAO25,HAO26=:HAO26,HAO27=:HAO27,HAO28=:HAO28,BZ=:BZ,SJ=:SJ,PJZ=:PJZ,HAO29=:HAO29,HAO30=:HAO30,HAO31=:HAO31,HAO32=:HAO32,HAO33=:HAO33,HAO34=:HAO34,HAO35=:HAO35,HAO36=:HAO36,HAO37=:HAO37,HAO38=:HAO38,HAO39=:HAO39,HAO40=:HAO40,HAO41=:HAO41,HAO42=:HAO42,HAO43=:HAO43,HAO44=:HAO44,HAO45=:HAO45,HAO46=:HAO46,HAO47=:HAO47,HAO48=:HAO48,HAO49=:HAO49,HAO50=:HAO50,CYS=:CYS,GYS=:GYS WHERE ROWID='" + RowId.Value + "'";
                updateCmd.Parameters.Add(":GAOLU", OracleType.Int32, 22).Value = (object)this.GAOLU ?? DBNull.Value;
                updateCmd.Parameters.Add(":BC", OracleType.VarChar, 10).Value = this.BC;
                updateCmd.Parameters.Add(":CC", OracleType.VarChar, 8).Value = this.CC;
                updateCmd.Parameters.Add(":HAO1", OracleType.Double, 22).Value = (object)this.HAO1 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO2", OracleType.DateTime, 7).Value = (object)this.HAO2 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO3", OracleType.Double, 22).Value = (object)this.HAO3 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO4", OracleType.Double, 22).Value = (object)this.HAO4 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO5", OracleType.Double, 22).Value = (object)this.HAO5 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO6", OracleType.Double, 22).Value = (object)this.HAO6 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO7", OracleType.Double, 22).Value = (object)this.HAO7 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO8", OracleType.Double, 22).Value = (object)this.HAO8 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO9", OracleType.Double, 22).Value = (object)this.HAO9 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO10", OracleType.Double, 22).Value = (object)this.HAO10 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO11", OracleType.Double, 22).Value = (object)this.HAO11 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO12", OracleType.Double, 22).Value = (object)this.HAO12 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO13", OracleType.Double, 22).Value = (object)this.HAO13 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO14", OracleType.Double, 22).Value = (object)this.HAO14 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO15", OracleType.Double, 22).Value = (object)this.HAO15 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO16", OracleType.Double, 22).Value = (object)this.HAO16 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO17", OracleType.Double, 22).Value = (object)this.HAO17 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO18", OracleType.Double, 22).Value = (object)this.HAO18 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO19", OracleType.Double, 22).Value = (object)this.HAO19 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO20", OracleType.Double, 22).Value = (object)this.HAO20 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO21", OracleType.Double, 22).Value = (object)this.HAO21 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO22", OracleType.Double, 22).Value = (object)this.HAO22 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO23", OracleType.Double, 22).Value = (object)this.HAO23 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO24", OracleType.Double, 22).Value = (object)this.HAO24 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO25", OracleType.Double, 22).Value = (object)this.HAO25 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO26", OracleType.Double, 22).Value = (object)this.HAO26 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO27", OracleType.Double, 22).Value = (object)this.HAO27 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO28", OracleType.Double, 22).Value = (object)this.HAO28 ?? DBNull.Value;
                updateCmd.Parameters.Add(":BZ", OracleType.VarChar, 20).Value = this.BZ;
                updateCmd.Parameters.Add(":SJ", OracleType.DateTime, 7).Value = (object)this.SJ ?? DBNull.Value;
                updateCmd.Parameters.Add(":PJZ", OracleType.Double, 22).Value = (object)this.PJZ ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO29", OracleType.Double, 22).Value = (object)this.HAO29 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO30", OracleType.Double, 22).Value = (object)this.HAO30 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO31", OracleType.Double, 22).Value = (object)this.HAO31 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO32", OracleType.Double, 22).Value = (object)this.HAO32 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO33", OracleType.Double, 22).Value = (object)this.HAO33 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO34", OracleType.Double, 22).Value = (object)this.HAO34 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO35", OracleType.Double, 22).Value = (object)this.HAO35 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO36", OracleType.Double, 22).Value = (object)this.HAO36 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO37", OracleType.Double, 22).Value = (object)this.HAO37 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO38", OracleType.Double, 22).Value = (object)this.HAO38 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO39", OracleType.Double, 22).Value = (object)this.HAO39 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO40", OracleType.Double, 22).Value = (object)this.HAO40 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO41", OracleType.Double, 22).Value = (object)this.HAO41 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO42", OracleType.Double, 22).Value = (object)this.HAO42 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO43", OracleType.Double, 22).Value = (object)this.HAO43 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO44", OracleType.Double, 22).Value = (object)this.HAO44 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO45", OracleType.Double, 22).Value = (object)this.HAO45 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO46", OracleType.Double, 22).Value = (object)this.HAO46 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO47", OracleType.Double, 22).Value = (object)this.HAO47 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO48", OracleType.Double, 22).Value = (object)this.HAO48 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO49", OracleType.Double, 22).Value = (object)this.HAO49 ?? DBNull.Value;
                updateCmd.Parameters.Add(":HAO50", OracleType.Double, 22).Value = (object)this.HAO50 ?? DBNull.Value;
                updateCmd.Parameters.Add(":CYS", OracleType.Double, 22).Value = (object)this.CYS ?? DBNull.Value;
                updateCmd.Parameters.Add(":GYS", OracleType.Double, 22).Value = (object)this.GYS ?? DBNull.Value;
                updateCmd.ExecuteNonQuery();
            }


        }
    }
    public partial class RblqTable : DbEntityTable<Rblq>, IFilter
    {
        private DateTime? loadRq = null;
        private int? loadGaolu = null;
        //Rblq,ltznConnectionString
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
            selectCmd.CommandText = "SELECT GAOLU,BC,CC,HAO1,HAO2,HAO3,HAO4,HAO5,HAO6,HAO7,HAO8,HAO9,HAO10,HAO11,HAO12,HAO13,HAO14,HAO15,HAO16,HAO17,HAO18,HAO19,HAO20,HAO21,HAO22,HAO23,HAO24,HAO25,HAO26,HAO27,HAO28,BZ,SJ,PJZ,HAO29,HAO30,HAO31,HAO32,HAO33,HAO34,HAO35,HAO36,HAO37,HAO38,HAO39,HAO40,HAO41,HAO42,HAO43,HAO44,HAO45,HAO46,HAO47,HAO48,HAO49,HAO50,CYS,GYS,ROWID FROM RBLQ WHERE TRUNC(SJ)=:RQ AND GAOLU=:GAOLU order by HAO2";
            selectCmd.Parameters.Add(":RQ", OracleType.DateTime).Value = rq;
            selectCmd.Parameters.Add(":GAOLU", OracleType.Int32).Value = gaolu;

            OracleDataReader dr = selectCmd.ExecuteReader();
            while (dr.Read())
            {
                Rblq item = new Rblq();
                item.BeginInit();
                if (dr.IsDBNull(0)) item.GAOLU = null; else item.GAOLU = dr.GetInt32(0);
                if (dr.IsDBNull(1)) item.BC = ""; else item.BC = dr.GetString(1);
                if (dr.IsDBNull(2)) item.CC = ""; else item.CC = dr.GetString(2);
                if (dr.IsDBNull(3)) item.HAO1 = null; else item.HAO1 = dr.GetDouble(3);
                if (dr.IsDBNull(4)) item.HAO2 = null; else item.HAO2 = dr.GetDateTime(4);
                if (dr.IsDBNull(5)) item.HAO3 = null; else item.HAO3 = dr.GetDouble(5);
                if (dr.IsDBNull(6)) item.HAO4 = null; else item.HAO4 = dr.GetDouble(6);
                if (dr.IsDBNull(7)) item.HAO5 = null; else item.HAO5 = dr.GetDouble(7);
                if (dr.IsDBNull(8)) item.HAO6 = null; else item.HAO6 = dr.GetDouble(8);
                if (dr.IsDBNull(9)) item.HAO7 = null; else item.HAO7 = dr.GetDouble(9);
                if (dr.IsDBNull(10)) item.HAO8 = null; else item.HAO8 = dr.GetDouble(10);
                if (dr.IsDBNull(11)) item.HAO9 = null; else item.HAO9 = dr.GetDouble(11);
                if (dr.IsDBNull(12)) item.HAO10 = null; else item.HAO10 = dr.GetDouble(12);
                if (dr.IsDBNull(13)) item.HAO11 = null; else item.HAO11 = dr.GetDouble(13);
                if (dr.IsDBNull(14)) item.HAO12 = null; else item.HAO12 = dr.GetDouble(14);
                if (dr.IsDBNull(15)) item.HAO13 = null; else item.HAO13 = dr.GetDouble(15);
                if (dr.IsDBNull(16)) item.HAO14 = null; else item.HAO14 = dr.GetDouble(16);
                if (dr.IsDBNull(17)) item.HAO15 = null; else item.HAO15 = dr.GetDouble(17);
                if (dr.IsDBNull(18)) item.HAO16 = null; else item.HAO16 = dr.GetDouble(18);
                if (dr.IsDBNull(19)) item.HAO17 = null; else item.HAO17 = dr.GetDouble(19);
                if (dr.IsDBNull(20)) item.HAO18 = null; else item.HAO18 = dr.GetDouble(20);
                if (dr.IsDBNull(21)) item.HAO19 = null; else item.HAO19 = dr.GetDouble(21);
                if (dr.IsDBNull(22)) item.HAO20 = null; else item.HAO20 = dr.GetDouble(22);
                if (dr.IsDBNull(23)) item.HAO21 = null; else item.HAO21 = dr.GetDouble(23);
                if (dr.IsDBNull(24)) item.HAO22 = null; else item.HAO22 = dr.GetDouble(24);
                if (dr.IsDBNull(25)) item.HAO23 = null; else item.HAO23 = dr.GetDouble(25);
                if (dr.IsDBNull(26)) item.HAO24 = null; else item.HAO24 = dr.GetDouble(26);
                if (dr.IsDBNull(27)) item.HAO25 = null; else item.HAO25 = dr.GetDouble(27);
                if (dr.IsDBNull(28)) item.HAO26 = null; else item.HAO26 = dr.GetDouble(28);
                if (dr.IsDBNull(29)) item.HAO27 = null; else item.HAO27 = dr.GetDouble(29);
                if (dr.IsDBNull(30)) item.HAO28 = null; else item.HAO28 = dr.GetDouble(30);
                if (dr.IsDBNull(31)) item.BZ = ""; else item.BZ = dr.GetString(31);
                if (dr.IsDBNull(32)) item.SJ = null; else item.SJ = dr.GetDateTime(32);
                if (dr.IsDBNull(33)) item.PJZ = null; else item.PJZ = dr.GetDouble(33);
                if (dr.IsDBNull(34)) item.HAO29 = null; else item.HAO29 = dr.GetDouble(34);
                if (dr.IsDBNull(35)) item.HAO30 = null; else item.HAO30 = dr.GetDouble(35);
                if (dr.IsDBNull(36)) item.HAO31 = null; else item.HAO31 = dr.GetDouble(36);
                if (dr.IsDBNull(37)) item.HAO32 = null; else item.HAO32 = dr.GetDouble(37);
                if (dr.IsDBNull(38)) item.HAO33 = null; else item.HAO33 = dr.GetDouble(38);
                if (dr.IsDBNull(39)) item.HAO34 = null; else item.HAO34 = dr.GetDouble(39);
                if (dr.IsDBNull(40)) item.HAO35 = null; else item.HAO35 = dr.GetDouble(40);
                if (dr.IsDBNull(41)) item.HAO36 = null; else item.HAO36 = dr.GetDouble(41);
                if (dr.IsDBNull(42)) item.HAO37 = null; else item.HAO37 = dr.GetDouble(42);
                if (dr.IsDBNull(43)) item.HAO38 = null; else item.HAO38 = dr.GetDouble(43);
                if (dr.IsDBNull(44)) item.HAO39 = null; else item.HAO39 = dr.GetDouble(44);
                if (dr.IsDBNull(45)) item.HAO40 = null; else item.HAO40 = dr.GetDouble(45);
                if (dr.IsDBNull(46)) item.HAO41 = null; else item.HAO41 = dr.GetDouble(46);
                if (dr.IsDBNull(47)) item.HAO42 = null; else item.HAO42 = dr.GetDouble(47);
                if (dr.IsDBNull(48)) item.HAO43 = null; else item.HAO43 = dr.GetDouble(48);
                if (dr.IsDBNull(49)) item.HAO44 = null; else item.HAO44 = dr.GetDouble(49);
                if (dr.IsDBNull(50)) item.HAO45 = null; else item.HAO45 = dr.GetDouble(50);
                if (dr.IsDBNull(51)) item.HAO46 = null; else item.HAO46 = dr.GetDouble(51);
                if (dr.IsDBNull(52)) item.HAO47 = null; else item.HAO47 = dr.GetDouble(52);
                if (dr.IsDBNull(53)) item.HAO48 = null; else item.HAO48 = dr.GetDouble(53);
                if (dr.IsDBNull(54)) item.HAO49 = null; else item.HAO49 = dr.GetDouble(54);
                if (dr.IsDBNull(55)) item.HAO50 = null; else item.HAO50 = dr.GetDouble(55);
                if (dr.IsDBNull(56)) item.CYS = null; else item.CYS = dr.GetDouble(56);
                if (dr.IsDBNull(57)) item.GYS = null; else item.GYS = dr.GetDouble(57);
                item.RowId = dr.GetString(58);

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
                  
                    if(item.SJ==null)
                        item.SJ=DateTime.Now.AddMinutes(i++);

                       item.SJ = new DateTime(loadRq.Value.Year, loadRq.Value.Month, loadRq.Value.Day, item.SJ.Value.Hour, item.SJ.Value.Minute, 0);
                   
                    item.GAOLU = loadGaolu.Value;
                }
            }
            base.Save();
        }
    }
}
