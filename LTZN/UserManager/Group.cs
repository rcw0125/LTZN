using System;
using System.Collections.Generic;
using System.Text;
using LTZN.Data;

namespace LTZN.UserManager
{
    public class Group : ObservableObject
    {
        private string _groupName = "";
        /// <summary>
        /// 用户组名
        /// </summary>
        public string GroupName
        {
            get { return _groupName; }
            set
            {
                if (_groupName != value)
                {
                    _groupName = value;
                    RaisePropertyChanged("GroupName");
                }
            }
        }

        private string _descr = "";
        /// <summary>
        /// 描述
        /// </summary>
        public string Descr
        {
            get { return _descr; }
            set
            {
                if (_descr != value)
                {
                    _descr = value;
                    RaisePropertyChanged("Descr");
                }
            }
        }

        private object sysnObj = new object();

        private List<string> userNames = new List<string>();

        private List<string> authorityNames = new List<string>();

        public bool HasAuthority(string authorityName)
        {
            return authorityName.Contains(authorityName);
        }

        public void AddAuthority(string authorityName)
        {
            lock(sysnObj)
            {
                if(!authorityNames.Contains(authorityName))
                authorityNames.Add(authorityName);
            }
        }

        public void RemoveAuthority(string authorityName)
        {
            lock (sysnObj)
            {
                if (authorityNames.Contains(authorityName))
                    authorityNames.Remove(authorityName);
            }
        }
    }
}
