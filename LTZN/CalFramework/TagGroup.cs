using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Text;

namespace LTZN.CalFramework
{
    public class TagGroup : List<string>
    {
        public string EID
        {
            get;
            set;
        }

        public string GroupName
        {
            get;
            set;
        }
    }

    public class TagOrganization
    {

        public string EID
        {
            get;
            set;
        }

        public CalModel model = null;

        public TagOrganization(CalModel model)
        {
            this.model = model;

        }

        public TagOrganization(CalModel model,string orgName)
        {
            this.model = model;
            this.OrganizationName = orgName;

        }
        public Dictionary<string, OrgTag> orgTags = new Dictionary<string, OrgTag>();

        private List<TagOrganization> subOrganizations = new List<TagOrganization>();

        public List<TagOrganization> SubOrganizations
        {
            get { return subOrganizations; }
        }

        public TagOrganization AddSubOrg(string orgName)
        {
            TagOrganization org = new TagOrganization(model, orgName);
            org.TagGroup = TagGroup;
            SubOrganizations.Add(org);
            return org;
        }

        public string OrganizationName
        {
            get;
            set;
        }

        public TagOrganization GetOrganization(string name)
        {
            if (this.OrganizationName == name)
                return this;
            foreach (var item in subOrganizations)
            {
               TagOrganization result=  item.GetOrganization(name);
               if (result != null)
                   return result;
            }
            return null;
        }

        private TagGroup tagGroup = null;

        public TagGroup TagGroup
        {
            get { return tagGroup; }
            set
            {

                if (tagGroup != value)
                {
                    orgTags.Clear();
                    tagGroup = value;
                    if (tagGroup != null)
                    {
                        foreach (string item in tagGroup)
                        {
                            OrgTag tag = new OrgTag(model, this, item);
                            orgTags.Add(item, tag);
                        }
                    }
                    foreach (TagOrganization subOrg in SubOrganizations)
                    {
                        subOrg.TagGroup = value;
                    }
                }

            }
        }

        public OrgTag  GetOrgTag(string groupItem)
        {
            return orgTags[groupItem];
        }

    }

    public class CalModel
    {

        private Calculate calculate = null;

        private List<TagOrganization> roots = new List<TagOrganization>();
      
        public CalModel()
        {
            calculate = new Calculate(this);
        }

        public string EID
        {
            get;
            set;
        }

        private string _modelName = "";
        public string ModelName
        {
            get { return _modelName; }
            set { _modelName = value; }
        }


        public TagOrganization AddRoot(string orgName)
        {
            TagOrganization org = new TagOrganization(this, orgName);
            roots.Add(org);
            return org;
        }

        public void AddTag(string tagName)
        {
            CalTag tag = new CalTag(this, tagName);
            CalTags.Add(tagName, tag);
        }

        public OrgTag GetOrgTag(string orgName, string groupItem)
        {
            foreach (TagOrganization root in roots)
            {
                if (root != null)
                {
                    TagOrganization org = root.GetOrganization(orgName);
                    return org.GetOrgTag(groupItem);
                }
            }
            return null;
        }

        private Dictionary<string, CalTag> _CalTags = new Dictionary<string, CalTag>();
        public Dictionary<string, CalTag> CalTags
        {
            get
            {
                return _CalTags;
            }
        }

        public void Calc()
        {
            foreach (CalTag tag in CalTags.Values)
            {
                tag.CalFinished = false;
            }

            foreach (CalTag tag in CalTags.Values)
            {
                tag.Cal();
            }
        }

        public double GetTagValue(string tagName)
        {
            if (tagName.Contains("."))
            {
                string[] nameItems = tagName.Split('.');
                if (nameItems.Length > 1)
                {
                    OrgTag orgtag = GetOrgTag(nameItems[0], nameItems[1]);
                    return orgtag.TagValue;
                }
            }
            else
                return CalTags[tagName].TagValue;

            return 0;
        }

        public double DoForma(string forma)
        {
            return Convert.ToDouble(calculate.DoCal(forma));
        }
    
    }

    public class OrgTag : CalTag
    {

        TagOrganization orgItem = null;
        string groupItem = "";

        public OrgTag(CalModel model, TagOrganization orgItem, string groupItem)
            : base(model)
        {
            this.orgItem = orgItem;
            this.groupItem = groupItem;
        }
        
        public override void Cal()
        {
            double total = 0;
            foreach (var item in orgItem.SubOrganizations)
            {
                total += item.GetOrgTag(groupItem).TagValue;
            }
            this.TagValue = total;
            base.Cal();
        }
    }


}
