using System.ComponentModel;
using System.Runtime.CompilerServices;
using Microsoft.Practices.EnterpriseLibrary.Logging;
using System;
using System.Text;

namespace WhatsDependentAndHow.ViewModels
{
    public class PropertyChangedViewModel : INotifyPropertyChanged
    {
        private StringBuilder _statusMessage = new StringBuilder();
        public string StatusMessage
        {
            get { return _statusMessage.ToString(); }
            set
            {
                _statusMessage.Append(value);
                OnPropertyChanged("StatusMessage");
            }
        }

        public void ClearStatusMessage()
        {
            _statusMessage.Clear();
        }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public void UpdateStatus(string message)
        {
            Logger.Write(message);
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                StatusMessage = string.Format("[{0} {1}]: {2}\n", DateTime.Now.ToShortDateString(), DateTime.Now.ToShortTimeString(), message);
            });
        }
    }
}
