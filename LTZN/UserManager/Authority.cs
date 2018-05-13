using System;
using System.Collections.Generic;
using System.Text;
using LTZN.Data;

namespace LTZN.UserManager
{
    public class Authority : ObservableObject
    {
        private string _authorityName = "";
        /// <summary>
        /// 用户组名
        /// </summary>
        public string AuthorityName
        {
            get { return _authorityName; }
            set
            {
                if (_authorityName != value)
                {
                    _authorityName = value;
                    RaisePropertyChanged("AuthorityName");
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
    }
}
