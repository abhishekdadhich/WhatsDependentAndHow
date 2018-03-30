using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Input;
using WhatsDependentAndHow.Commands;
using Excel = Microsoft.Office.Interop.Excel;

namespace WhatsDependentAndHow
{
    public class MainWindowViewModel : INotifyPropertyChanged, IDisposable, IDataErrorInfo
    {
        public bool IsExcelFileInfoLoaded { get; set; }

        private bool _isBusy;
        public bool IsBusy
        {
            get { return _isBusy; }
            set
            {
                _isBusy = value;
                OnPropertyChanged("IsBusy");
            }
        }

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

        private string _selectedSheetName = string.Empty;
        public string SelectedSheetName
        {
            get { return _selectedSheetName; }
            set
            {
                _selectedSheetName = value;
                OnPropertyChanged("SelectedSheetName");
            }
        }

        private bool _isCellAddressValid = false;
        public bool IsCellAddressValid
        {
            get { return _isCellAddressValid; }
            set
            {
                _isCellAddressValid = value;
                OnPropertyChanged("IsCellAddressValid");
            }
        }

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

        private ObservableCollection<string> _worksheetNames = new ObservableCollection<string>();
        public ObservableCollection<string> WorkSheetNames
        {
            get { return _worksheetNames; }
            set
            {
                _worksheetNames = value;
                OnPropertyChanged("WorkSheetNames");
            }
        }

        private Excel.Application _xlApp = null;
        public Excel.Application XlApp
        {
            get { return _xlApp; }
            set
            {
                _xlApp = value;
                OnPropertyChanged("XlApp");
            }
        }

        private Excel.Workbook _xlWorkBook = null;
        public Excel.Workbook XlWorkBook
        {
            get { return _xlWorkBook; }
            set
            {
                _xlWorkBook = value;
                OnPropertyChanged("XlWorkBook");
            }
        }

        private CellObject _workBookRootCellObject = new CellObject();
        public CellObject WorkBookRootCellObject
        {
            get { return _workBookRootCellObject; }
            set
            {
                _workBookRootCellObject = value;
                OnPropertyChanged("WorkBookRootCellObject");
            }
        }

        private string _statusMessage = "Awaiting User Input...";
        public string StatusMessage
        {
            get { return _statusMessage; }
            set
            {
                _statusMessage = value;
                OnPropertyChanged("StatusMessage");
            }
        }

        private ButtonFileSelectorCommand _buttonFileSelector;
        private ButtonOutputPathSelectorCommand _buttonOutputPathSelector;
        private ButtonGenerateTreeCommand _buttonGenerateTree;
        private ButtonCloseApplicationCommand _buttonCloseApplicationCommand;

        public MainWindowViewModel()
        {
            _buttonFileSelector = new ButtonFileSelectorCommand(this);
            _buttonOutputPathSelector = new ButtonOutputPathSelectorCommand(this);
            _buttonGenerateTree = new ButtonGenerateTreeCommand(this);
            _buttonCloseApplicationCommand = new ButtonCloseApplicationCommand(this);
        }

        public ICommand ButtonFileSelectorClickCommand
        {
            get { return _buttonFileSelector; }
        }

        public ICommand ButtonOutputPathSelectorClickCommand
        {
            get { return _buttonOutputPathSelector; }
        }

        public ICommand ButtonGenerateTreeClickCommand
        {
            get { return _buttonGenerateTree; }
        }

        public ICommand ButtonCloseApplicationClickCommand
        {
            get { return _buttonCloseApplicationCommand; }
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
            {
                IsCellAddressValid = false;
                return "Invalid Cell Address: " + CellAddress;
            }
            else
            {
                IsCellAddressValid = true;
                return string.Empty;
            }
        }

        public void ClearControls()
        {
            ExcelFileDetails = string.Empty;
            OutputFilePath = string.Empty;
            CellAddress = string.Empty;
            StatusMessage = string.Empty;
            _worksheetNames.Clear();
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
