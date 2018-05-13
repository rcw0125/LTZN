using System;
using System.Collections.Generic;
using System.Text;

namespace LTZN.CalFramework
{

    public class TagGroupRepository
    {

        public TagGroup CreateTagGroup(string groupName)
        {

            Guid id = Guid.NewGuid();
            id.ToString();
            TagGroup tagGroup = new TagGroup();
            tagGroup.EID = id.ToString();
            tagGroup.GroupName = groupName;
            
            Calf_Taggroup a = new Calf_Taggroup();
            Calf_Taggroup_Item item = new Calf_Taggroup_Item();
            

            return tagGroup;
        }

        public List<TagGroup> GetALlGroup()
        {
            return new List<TagGroup>();
        }

        public void Remove(TagGroup tagGroup)
        {

        }

        public void Update(TagGroup tagGroup)
        {

        }
   
    }
}
