using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;

namespace LTZN.Repository
{
    public class RepositoryBase 
    {
        protected OracleConnection conn = new OracleConnection();

        public RepositoryBase()
        {
            conn.ConnectionString = Properties.Settings.Default.ltznConnectionString;
        }

    }
}
