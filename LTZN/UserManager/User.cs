using System;
using System.Collections.Generic;
using System.Text;
using LTZN.Data;

namespace LTZN.UserManager
{
    public class User : ObservableObject
    {
        private string _userName = "";
        /// <summary>
        /// 用户名
        /// </summary>
        public string UserName
        {
            get { return _userName; }
            set
            {
                if (_userName != value)
                {
                    _userName = value;
                    RaisePropertyChanged("UserName");
                }
            }
        }

        private string _password = "";
        /// <summary>
        /// 密码
        /// </summary>
        public string Password
        {
            get { return _password; }
            set
            {
                if (_password != value)
                {
                    _password = value;
                    RaisePropertyChanged("Password");
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


        private DateTime _createTime = DateTime.Now;
        /// <summary>
        /// 创建时间
        /// </summary>
        public DateTime CreateTime
        {
            get { return _createTime; }
            set
            {
                if (_createTime != value)
                {
                    _createTime = value;
                    RaisePropertyChanged("CreateTime");
                }
            }
        }


        private DateTime _lastLoginTime;
        /// <summary>
        /// 最后一次登陆时间
        /// </summary>
        public DateTime LastLoginTime
        {
            get { return _lastLoginTime; }
            set
            {
                if (_lastLoginTime != value)
                {
                    _lastLoginTime = value;
                    RaisePropertyChanged("LastLoginTime");
                }
            }
        }

        private string _lastLoginIp;
        /// <summary>
        /// 最后一次登陆IP
        /// </summary>
        public string LastLoginIp
        {
            get { return _lastLoginIp; }
            set
            {
                if (_lastLoginIp != value)
                {
                    _lastLoginIp = value;
                    RaisePropertyChanged("LastLoginIp");
                }
            }
        }

        private List<string> groupNames = new List<string>();

        /// <summary>
        /// 是否属于用户组
        /// </summary>
        /// <param name="groupName"></param>
        /// <returns></returns>
        public bool IsGroup(string groupName)
        {
            return groupNames.Contains(groupName);
        }

        private List<string> authorityNames = new List<string>();
        /// <summary>
        /// 是否有权限
        /// </summary>
        /// <param name="authorityName"></param>
        /// <returns></returns>
        public bool HasAuthority(string authorityName)
        {
            return authorityName.Contains(authorityName);
        }
    }
}
