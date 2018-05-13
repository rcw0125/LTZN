using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.Query
{
    public class QueryEntity : List<EntityField>
    {

        private string _QueryEntityName = "";
        public string QueryEntityName
        {
            get
            {
                return _QueryEntityName;
            }
        }

        public QueryEntity(string name)
        {
            this._QueryEntityName = name;
        }


        public EntityField GetByDisplay(string display)
        {
            foreach (var item in this)
            {
                if (item.FieldDispaly == display)
                    return item;
            }

            return null;
        }

    }
}
