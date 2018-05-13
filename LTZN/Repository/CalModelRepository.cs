using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;
using LTZN.CalFramework;

namespace LTZN.Repository
{
    public class CalModelRepository : RepositoryBase
    {
        public List<CalModel> GetAllModels()
        {
            Dictionary<string, CalModel> allModels = new Dictionary<string, CalModel>();
            Dictionary<string, CalOrg> allOrgs = new Dictionary<string, CalOrg>();

            conn.Open();
            OracleCommand selectModelCmd = new OracleCommand();
            selectModelCmd.Connection = conn;
            selectModelCmd.CommandText = "Select EID,MODELNAME From CALF_CALMODEL order by MODELNAME";
            OracleDataReader drModel = selectModelCmd.ExecuteReader();
            while (drModel.Read())
            {
                CalModel model = new CalModel();
                model.EID = drModel.GetString(0);
                model.ModelName = drModel.IsDBNull(1) ? "" : drModel.GetString(1);
                allModels.Add(model.EID, model);
                model.ReSet();
            }
            drModel.Close();

            OracleCommand selectOrgCmd = new OracleCommand();
            selectOrgCmd.Connection = conn;
            selectOrgCmd.CommandText = "Select EID,CALMODELID,PARENTID,ORGNAME From CALF_CALORG order by orgname";
            OracleDataReader drOrg = selectOrgCmd.ExecuteReader();
            while (drOrg.Read())
            {
                CalOrg org = new CalOrg();
                org.EID = drOrg.GetString(0);
                if (allOrgs.ContainsKey(org.EID))
                    org = allOrgs[org.EID];
                else
                    allOrgs.Add(org.EID, org);
                CalModel myModel = allModels[drOrg.GetString(1)];
                org.MyModel = myModel;
                myModel.AllCalOrgs.Add(org);
                if (!drOrg.IsDBNull(2))
                {
                    string parentOrgID = drOrg.GetString(2);
                    if (allOrgs.ContainsKey(parentOrgID))
                    {
                        org.ParentCalOrg = allOrgs[parentOrgID];
                        org.ParentCalOrg.ChildCalOrgs.Add(org);
                    }
                    else
                    {
                        CalOrg parentOrg = new CalOrg();
                        parentOrg.EID = parentOrgID;
                        org.ParentCalOrg = parentOrg;
                        parentOrg.ChildCalOrgs.Add(org);
                        allOrgs.Add(parentOrgID, parentOrg);
                    }
                }
                else
                {
                    myModel.RootOrgs.Add(org);
                }
                org.OrgName = drOrg.IsDBNull(3) ? "" : drOrg.GetString(3);

                org.ReSet();
            }
            drOrg.Close();

            OracleCommand selectTagCmd = new OracleCommand();
            selectTagCmd.Connection = conn;
            selectTagCmd.CommandText = "Select EID,CALMODELID,CALORGID,TAGNAME,FORMA,DEC From CALF_CALTAG order by CALMODELID,CALORGID,IDX";
            OracleDataReader drTag = selectTagCmd.ExecuteReader();
            while (drTag.Read())
            {
                CalTag tag = new CalTag();
                tag.EID = drTag.GetString(0);
                CalModel myModel = allModels[drTag.GetString(1)];
                myModel.AllCalTags.Add(tag);
                tag.MyModel = myModel;

                if (!drTag.IsDBNull(2))
                {
                    string parentOrgID = drTag.GetString(2);
                    allOrgs[parentOrgID].CalTags.Add(tag);
                    tag.MyCalOrg = allOrgs[parentOrgID];

                }
                else
                {
                    myModel.RootTags.Add(tag);
                }

                tag.TagName = drTag.IsDBNull(3) ? "" : drTag.GetString(3);
                tag.Forma = drTag.IsDBNull(4) ? "" : drTag.GetString(4);
                if (!drTag.IsDBNull(5)) tag.Dec = drTag.GetInt32(5);
                tag.ReSet();

            }
            drTag.Close();

            conn.Close();

            List<CalModel> result = new List<CalModel>();
            result.AddRange(allModels.Values);
            return result;
        }

        public CalModel GetModel(string modelName)
        {
            CalModel model = null;

            Dictionary<string, CalOrg> allOrgs = new Dictionary<string, CalOrg>();

            conn.Open();
            OracleCommand selectModelCmd = new OracleCommand();
            selectModelCmd.Connection = conn;
            selectModelCmd.CommandText = "Select EID,MODELNAME From CALF_CALMODEL Where MODELNAME='" + modelName + "'";
            OracleDataReader drModel = selectModelCmd.ExecuteReader();
            if (drModel.Read())
            {
                model = new CalModel();
                model.EID = drModel.GetString(0);
                model.ModelName = drModel.IsDBNull(1) ? "" : drModel.GetString(1);
                model.ReSet();
            }
            drModel.Close();

            if (model != null)
            {
                OracleCommand selectOrgCmd = new OracleCommand();
                selectOrgCmd.Connection = conn;
                selectOrgCmd.CommandText = "Select EID,CALMODELID,PARENTID,ORGNAME From CALF_CALORG Where CALMODELID='" + model.EID + "' ORDER By ORGNAME";
                OracleDataReader drOrg = selectOrgCmd.ExecuteReader();
                while (drOrg.Read())
                {
                    CalOrg org = new CalOrg();
                    org.EID = drOrg.GetString(0);
                    if (allOrgs.ContainsKey(org.EID))
                        org = allOrgs[org.EID];
                    else
                        allOrgs.Add(org.EID, org);  
                    org.MyModel = model;
                    model.AllCalOrgs.Add(org);
                    if (!drOrg.IsDBNull(2))
                    {
                        string parentOrgID = drOrg.GetString(2);
                        if (allOrgs.ContainsKey(parentOrgID))
                        {
                            org.ParentCalOrg = allOrgs[parentOrgID];
                            org.ParentCalOrg.ChildCalOrgs.Add(org);
                        }
                        else
                        {
                            CalOrg parentOrg = new CalOrg();
                            parentOrg.EID = parentOrgID;
                            org.ParentCalOrg = parentOrg;
                            parentOrg.ChildCalOrgs.Add(org);
                            allOrgs.Add(parentOrgID, parentOrg);
                        }
                    }
                    else
                    {
                        model.RootOrgs.Add(org);
                    }
                    org.OrgName = drOrg.IsDBNull(3) ? "" : drOrg.GetString(3);
                    org.ReSet();
                }
                drOrg.Close();

                OracleCommand selectTagCmd = new OracleCommand();
                selectTagCmd.Connection = conn;
                selectTagCmd.CommandText = "Select EID,CALMODELID,CALORGID,TAGNAME,FORMA,DEC From CALF_CALTAG Where CALMODELID='" + model.EID + "' order by IDX";
                OracleDataReader drTag = selectTagCmd.ExecuteReader();
                while (drTag.Read())
                {
                    CalTag tag = new CalTag();
                    tag.EID = drTag.GetString(0);
                    model.AllCalTags.Add(tag);
                    tag.MyModel = model;

                    if (!drTag.IsDBNull(2) && allOrgs.ContainsKey(drTag.GetString(2)))
                    {
                        string parentOrgID = drTag.GetString(2);
                        allOrgs[parentOrgID].CalTags.Add(tag);
                        tag.MyCalOrg = allOrgs[parentOrgID];

                    }
                    else
                    {
                        model.RootTags.Add(tag);
                    }

                    tag.TagName = drTag.IsDBNull(3) ? "" : drTag.GetString(3);
                    tag.Forma = drTag.IsDBNull(4) ? "" : drTag.GetString(4);
                    if (!drTag.IsDBNull(5)) tag.Dec = drTag.GetInt32(5);
                    tag.ReSet();

                }
                drTag.Close();
               
            }
            conn.Close();
            return model;
        }

        public void Save(CalModel model)
        {
            conn.Open();
            OracleTransaction trans = conn.BeginTransaction();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;
            cmd.Transaction = trans;
            cmd.CommandText = "Delete From CALF_CALMODEL where EID='" + model.EID + "'";
            cmd.ExecuteNonQuery();
            cmd.CommandText = "Delete From CALF_CALORG where CALMODELID='" + model.EID + "'";
            cmd.ExecuteNonQuery();
            cmd.CommandText = "Delete From CALF_CALTAG where CALMODELID='" + model.EID + "'";
            cmd.ExecuteNonQuery();

            cmd.CommandText = "Insert Into CALF_CALMODEL(EID,MODELNAME) Values('" + model.EID + "','" + model.ModelName + "')";
            cmd.ExecuteNonQuery();
            foreach (var item in model.AllCalOrgs)
            {
                cmd.CommandText = "Insert Into CALF_CALORG(EID,CALMODELID,PARENTID,ORGNAME) Values(:EID,:CALMODELID,:PARENTID,:ORGNAME)";
                cmd.Parameters.Clear();
                cmd.Parameters.Add(":EID", OracleType.Char, 36).Value = item.EID;
                cmd.Parameters.Add(":CALMODELID", OracleType.Char, 36).Value = model.EID;
                cmd.Parameters.Add(":PARENTID", OracleType.Char, 36).Value = item.ParentCalOrg == null ? DBNull.Value : (object)item.ParentCalOrg.EID;
                cmd.Parameters.Add(":ORGNAME", OracleType.VarChar, 200).Value = item.OrgName ?? "";
                cmd.ExecuteNonQuery();
            }
            //int i = 1;
            foreach (var item in model.AllCalTags)
            {

                cmd.CommandText = "Insert Into CALF_CALTAG(EID,CALMODELID,CALORGID,TAGNAME,FORMA,DEC,IDX) Values(:EID,:CALMODELID,:CALORGID,:TAGNAME,:FORMA,:DEC,:IDX)";
                cmd.Parameters.Clear();
                cmd.Parameters.Add(":EID", OracleType.Char, 36).Value = item.EID;
                cmd.Parameters.Add(":CALMODELID", OracleType.Char, 36).Value = model.EID;
                cmd.Parameters.Add(":CALORGID", OracleType.Char, 36).Value = item.MyCalOrg == null ? DBNull.Value : (object)item.MyCalOrg.EID;
                cmd.Parameters.Add(":TAGNAME", OracleType.VarChar, 200).Value = item.TagName ?? "";
                cmd.Parameters.Add(":FORMA", OracleType.VarChar, 2000).Value = item.Forma ?? "";
                cmd.Parameters.Add(":DEC", OracleType.Int32).Value = (object)item.Dec ?? DBNull.Value;
                cmd.Parameters.Add(":IDX", OracleType.Int32).Value = item.Idx;
                cmd.ExecuteNonQuery();
            }
            trans.Commit();
            conn.Close();
        }

        public void Delete(CalModel model)
        {
            conn.Open();
            OracleTransaction trans = conn.BeginTransaction();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;
            cmd.Transaction = trans;
            cmd.CommandText = "Delete From CALF_CALMODEL where EID='" + model.EID + "'";
            cmd.ExecuteNonQuery();
            cmd.CommandText = "Delete From CALF_CALORG where CALMODELID='" + model.EID + "'";
            cmd.ExecuteNonQuery();
            cmd.CommandText = "Delete From CALF_CALTAG where CALMODELID='" + model.EID + "'";
            cmd.ExecuteNonQuery();
            trans.Commit();
            conn.Close();
        }
    }
}
