using System;
using System.Collections.Generic;
using System.Text;
using System.Data.OracleClient;
using LTZN.CalFramework;
using LTZN.Repository;

namespace LTZN.质量日报
{
    public class 质量日报 : QueryModel
    {
        public 质量日报()
            : base("质量日报")
        {

        }

        public void getDataBy(DateTime rq)
        {
            conn.Open();
            OracleCommand cmd = new OracleCommand();
            cmd.Connection = conn;

            cmd.CommandText = "select * from ddyl where trunc(sj)=:rq";
            cmd.Parameters.Add(":rq", OracleType.DateTime).Value = rq;

            OracleDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                if (!dr.IsDBNull(4)) this.SetValue("含铁原料成分(一).TFe", dr.GetDouble(4));
                if (!dr.IsDBNull(5)) this.SetValue("含铁原料成分(一).SiO2", dr.GetDouble(5));
                if (!dr.IsDBNull(6)) this.SetValue("含铁原料成分(一).CaO", dr.GetDouble(6));
                if (!dr.IsDBNull(7)) this.SetValue("含铁原料成分(一).FeO", dr.GetDouble(7));
                if (!dr.IsDBNull(8)) this.SetValue("含铁原料成分(一).MgO", dr.GetDouble(8));
                if (!dr.IsDBNull(9)) this.SetValue("含铁原料成分(一).S", dr.GetDouble(9));
                if (!dr.IsDBNull(10)) this.SetValue("含铁原料成分(一).Al2O3", dr.GetDouble(10));
                if (!dr.IsDBNull(14)) this.SetValue("含铁原料成分(一).TiO2", dr.GetDouble(14));
                if (!dr.IsDBNull(15)) this.SetValue("含铁原料成分(一).P", dr.GetDouble(15));
                if (!dr.IsDBNull(16)) this.SetValue("含铁原料成分(一).MnO", dr.GetDouble(16));
                if (!dr.IsDBNull(17)) this.SetValue("含铁原料成分(一).Cr", dr.GetDouble(17));
                if (!dr.IsDBNull(18)) this.SetValue("含铁原料成分(一).Pb", dr.GetDouble(18));
                if (!dr.IsDBNull(19)) this.SetValue("含铁原料成分(一).Zn", dr.GetDouble(19));
                if (!dr.IsDBNull(20)) this.SetValue("含铁原料成分(一).JJS", dr.GetDouble(20));
                if (!dr.IsDBNull(21)) this.SetValue("含铁原料成分(一).V2O5", dr.GetDouble(21));
                if (!dr.IsDBNull(22)) this.SetValue("含铁原料成分(一).Cu", dr.GetDouble(22));
                if (!dr.IsDBNull(23)) this.SetValue("含铁原料成分(一).Mo", dr.GetDouble(23));
                if (!dr.IsDBNull(24)) this.SetValue("含铁原料成分(一).Ni", dr.GetDouble(24));
                if (!dr.IsDBNull(25)) this.SetValue("含铁原料成分(一).Sn", dr.GetDouble(25));
                if (!dr.IsDBNull(26)) this.SetValue("含铁原料成分(一).Sb", dr.GetDouble(26));
                if (!dr.IsDBNull(27)) this.SetValue("含铁原料成分(一).As", dr.GetDouble(27));

            }
            dr.Close();
            conn.Close();
        
        
        
        
        
        
        
        
        
        
        }









    }
}
