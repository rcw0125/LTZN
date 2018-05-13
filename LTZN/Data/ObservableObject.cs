using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;

namespace LTZN.Data
{
    public class ObservableObject : INotifyPropertyChanged
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
            
            _entityState = EntityState.Modified;
        }

        private EntityState _entityState=EntityState.Unmodified;

        public EntityState EntityState
        {
            get { return _entityState; }
        }

        public void ReSet()
        {
            if (_entityState != EntityState.Unmodified)
            {
                _entityState = EntityState.Unmodified;

                var handler = PropertyChanged;

                if (handler != null)
                {
                    handler(this, new PropertyChangedEventArgs("EntityState"));
                }
            }
        }
        
    }

    public enum EntityState
    {
        Modified, //已修改
        Unmodified //未修改
    }
}
