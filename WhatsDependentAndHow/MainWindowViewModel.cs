using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace WhatsDependentAndHow
{
    public class MainWindowViewModel : INotifyPropertyChanged, IDisposable
    {
        public bool IsExcelFileInfoLoaded { get; set; }

        private string _excelFileDetails = string.Empty;
        public string ExcelFileDetails
        {
            get { return _excelFileDetails; }
            set
            {
                string inputValue = value.Trim();

                if (_excelFileDetails == inputValue)
                    return;

                _excelFileDetails = inputValue;

                IsExcelFileInfoLoaded = (_excelFileDetails.Length > 0) ? true : false;

                OnPropertyChanged("ExcelFileDetails");
            }
        }

        public bool IsOutputFilePathAvailable = false;
        private string _outputFilePath = string.Empty;
        public string OutputFilePath
        {
            get { return _outputFilePath; }
            set
            {
                string inputValue = value.Trim();
                if (_outputFilePath == inputValue)
                    return;

                _outputFilePath = value;

                IsOutputFilePathAvailable = (_outputFilePath.Length > 0) ? true : false;

                OnPropertyChanged("OutputFilePath");
            }
        }

        private ButtonFileSelectorCommand _buttonFileSelector;
        private ButtonOutputPathSelectorCommand _buttonOutputPathSelector;

        public MainWindowViewModel()
        {
            _buttonFileSelector = new ButtonFileSelectorCommand(this);
            _buttonOutputPathSelector = new ButtonOutputPathSelectorCommand(this);
        }

        public ICommand ButtonFileSelectorClickCommand
        {
            get { return _buttonFileSelector; }
        }

        public ICommand ButtonOutputPathSelectorClickCommand
        {
            get { return _buttonOutputPathSelector; }
        }

        #region INotifyPropertyChanged Support

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        #endregion

        #region IDisposable Support
        private bool disposedValue = false; // To detect redundant calls

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects).
                }

                disposedValue = true;
            }
        }


        public void Dispose()
        {
            Dispose(true);
        }
        #endregion
    }
}
