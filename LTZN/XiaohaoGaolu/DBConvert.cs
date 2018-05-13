using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;

namespace LTZN.XiaohaoGaolu
{
    class DBConvert
    {
        public static OracleParameter CreateOracleParameter(string name, Double? value)
        {
            OracleParameter p = new OracleParameter(name, OracleType.Double);
            if (value.HasValue)
            {
                p.Value = value.Value;
            }
            else
            {
                p.Value = System.DBNull.Value;
            }
            return p;
        }
        public static OracleParameter CreateOracleParameter(string name, int? value)
        {
            OracleParameter p = new OracleParameter(name, OracleType.Int32);
            if (value.HasValue)
            {
                p.Value = value.Value;
            }
            else
            {
                p.Value = System.DBNull.Value;
            }
            return p;
        }
        public static OracleParameter CreateOracleParameter(string name, int value)
        {
            OracleParameter p = new OracleParameter(name, OracleType.Int32);
            p.Value = value;
            return p;
        }
        public static OracleParameter CreateOracleParameter(string name, DateTime? value)
        {
            OracleParameter p = new OracleParameter(name, OracleType.DateTime);
            if (value.HasValue)
            {
                p.Value = value.Value;
            }
            else
            {
                p.Value = System.DBNull.Value;
            }
            return p;
        }
        public static OracleParameter CreateOracleParameter(string name, DateTime value)
        {
            OracleParameter p = new OracleParameter(name, OracleType.DateTime);
            p.Value = value;
            return p;
        }
        public static OracleParameter CreateOracleParameter(string name, string value)
        {
            OracleParameter p = new OracleParameter(name, OracleType.VarChar);
            if (value!=null)
            {
                p.Value = value;
            }
            else
            {
                p.Value = System.DBNull.Value;
            }
            return p;
        }
    }
}
