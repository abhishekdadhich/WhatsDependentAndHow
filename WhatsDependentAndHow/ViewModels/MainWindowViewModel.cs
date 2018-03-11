using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Input;

namespace WhatsDependentAndHow
{
    public class MainWindowViewModel : INotifyPropertyChanged, IDisposable, IDataErrorInfo
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
                OnPropertyChanged("IsExcelFileInfoLoaded");
            }
        }

        public bool IsOutputFilePathAvailable { get; set; }
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
                OnPropertyChanged("IsOutputFilePathAvailable");
            }
        }

        public bool IsCellAddressValid = true;
        private string _cellAddress = string.Empty;
        public string CellAddress
        {
            get { return _cellAddress; }
            set
            {
                string inputValue = value.Trim().ToUpper();

                if (inputValue == _cellAddress)
                    return;

                _cellAddress = inputValue;

                OnPropertyChanged("CellAddress");
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

        public string Error
        {
            get
            {
                return null;
            }
        }

        public string this[string columnName]
        {
            get
            {
                switch(columnName)
                {
                    case "CellAddress":
                        return ValidateCellAddress();
                }

                return string.Empty;
            }
        }

        private string ValidateCellAddress()
        {
            if (CellAddress.Length == 0)
                return string.Empty;

            string regExValidCell = @"^([A-Z]+)(\d+)$";

            Regex regex = new Regex(regExValidCell);
            if (!regex.IsMatch(CellAddress))
                return "Invalid Cell Address: " + CellAddress;
            else
                return string.Empty;
        }

        public void ClearControls()
        {
            ExcelFileDetails = string.Empty;
            OutputFilePath = string.Empty;
            CellAddress = string.Empty;
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
