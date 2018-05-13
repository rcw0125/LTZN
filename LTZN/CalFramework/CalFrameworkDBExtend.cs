using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.CalFramework
{
    public partial class Calf_Taggroup
    {
        public override string ToString()
        {
            return this.ENAME;
        }
    }

    public partial class Calf_Taggroup_Item
    {
        public override string ToString()
        {
            return this.TAGNAME;
        }
    }

    public partial class Calf_Tagorganization
    {
        private Calf_Taggroup _Calf_Taggroup = null;

        public Calf_Taggroup Calf_Taggroup
        {
            set
            {   _Calf_Taggroup = value;
                if (TAGGROUPID != _Calf_Taggroup.EID)
                {
                    Calf_CaltagTable.Clear();
                    foreach (var item in _Calf_Taggroup.Calf_Taggroup_ItemTable)
                    {
                        Calf_Caltag tag = new Calf_Caltag();
                        tag.EID = Guid.NewGuid().ToString();
                        tag.TAGNAME = this.ORGNAME + "." + item.TAGNAME;
                        tag.CALMODELID = this.CALMODELID;
                        tag.TAGGROUPID = _Calf_Taggroup.EID;
                        tag.TAGORGANIZATIONID = this.EID;
                        Calf_CaltagTable.Add(tag);
                    }
                }
            }
            get
            {
                if (TAGGROUPID != "" && _Calf_Taggroup == null)
                {
                    _Calf_Taggroup = Calf_Taggroup.LoadByPk(TAGGROUPID);
                }
                return _Calf_Taggroup;
            }
        }
    
    }
}
