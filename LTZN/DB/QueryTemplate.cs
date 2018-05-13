using System;
using System.ComponentModel;
using System.Data.OracleClient;

namespace ZHCDB
{
    public partial class Querytemplate : DbEntity
    {
        private string _ENTITYNAME = "";
        private string _TEMPLATENAME = "";
        private string _XMLTEXT = "";
        //ConStr 



        public string ENTITYNAME
        {
            get
            {
                return this._ENTITYNAME;
            }
            set
            {
                if (!_ENTITYNAME.Equals(value))
                {
                    _ENTITYNAME = value;
                    RaisePropertyChanged("ENTITYNAME", true);
                }

            }
        }
        public string TEMPLATENAME
        {
            get
            {
                return this._TEMPLATENAME;
            }
            set
            {
                if (!_TEMPLATENAME.Equals(value))
                {
                    _TEMPLATENAME = value;
                    RaisePropertyChanged("TEMPLATENAME", true);
                }

            }
        }
        public string XMLTEXT
        {
            get
            {
                return this._XMLTEXT;
            }
            set
            {
                if (!_XMLTEXT.Equals(value))
                {
                    _XMLTEXT = value;
                    RaisePropertyChanged("XMLTEXT", true);
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
                insertCmd.CommandText = "INSERT INTO QUERYTEMPLATE(ENTITYNAME,TEMPLATENAME,XMLTEXT) VALUES(:ENTITYNAME,:TEMPLATENAME,:XMLTEXT)";
                insertCmd.Parameters.Add(":ENTITYNAME", OracleType.VarChar, 200).Value = this.ENTITYNAME;
                insertCmd.Parameters.Add(":TEMPLATENAME", OracleType.VarChar, 200).Value = this.TEMPLATENAME;
                insertCmd.Parameters.Add(":XMLTEXT", OracleType.VarChar, 2000).Value = this.XMLTEXT;

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
                deleteCmd.CommandText = "DELETE FROM QUERYTEMPLATE WHERE ROWID='" + RowId.Value + "'";

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
                updateCmd.CommandText = "UPDATE QUERYTEMPLATE SET ENTITYNAME=:ENTITYNAME,TEMPLATENAME=:TEMPLATENAME,XMLTEXT=:XMLTEXT WHERE ROWID='" + RowId.Value + "'";
                updateCmd.Parameters.Add(":ENTITYNAME", OracleType.VarChar, 200).Value = this.ENTITYNAME;
                updateCmd.Parameters.Add(":TEMPLATENAME", OracleType.VarChar, 200).Value = this.TEMPLATENAME;
                updateCmd.Parameters.Add(":XMLTEXT", OracleType.VarChar, 2000).Value = this.XMLTEXT;

                updateCmd.ExecuteNonQuery();
            }


        }
    }
    public partial class QuerytemplateTable : DbEntityTable<Querytemplate>, IFilter
    {
        //Querytemplate,ltznConnectionString
        public void LoadByEntityName(string entityName)
        {
            //ConStr, EntiyType, SelectSql, SelectPara, SelectParaVal
            beginAdjustData = true;
            this.Clear();
            hideList.Clear();
            OracleConnection con = new OracleConnection(LTZN.Properties.Settings.Default.ltznConnectionString);
            con.Open();
            OracleCommand selectCmd = new OracleCommand();
            selectCmd.Connection = con;
            selectCmd.CommandText = "SELECT ENTITYNAME,TEMPLATENAME,XMLTEXT,ROWID FROM QUERYTEMPLATE WHERE ENTITYNAME=:ENTITYNAME";
            selectCmd.Parameters.Add(":ENTITYNAME", OracleType.VarChar).Value = entityName;

            OracleDataReader dr = selectCmd.ExecuteReader();
            while (dr.Read())
            {
                Querytemplate item = new Querytemplate();
                item.BeginInit();
                if (dr.IsDBNull(0)) item.ENTITYNAME = ""; else item.ENTITYNAME = dr.GetString(0);
                if (dr.IsDBNull(1)) item.TEMPLATENAME = ""; else item.TEMPLATENAME = dr.GetString(1);
                if (dr.IsDBNull(2)) item.XMLTEXT = ""; else item.XMLTEXT = dr.GetString(2);
                item.RowId = dr.GetString(3);

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


        public Querytemplate GetByTemplateName(string templateName)
        {
            foreach (var item in this)
            {
                if (item.TEMPLATENAME == templateName)
                {
                    return item;
                }
            }
            return null;
        }
    }
}
