using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;

namespace LTZN.Data
{
    public class DataObject : INotifyPropertyChanged
    {
        /// <summary>
        /// 属性值改变时发生.
        /// </summary>
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void RaisePropertyChanged(string propertyName)
        {
            var handler = PropertyChanged;

            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
