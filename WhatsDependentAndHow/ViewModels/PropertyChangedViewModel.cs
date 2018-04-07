using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace WhatsDependentAndHow.ViewModels
{
    public class PropertyChangedViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
