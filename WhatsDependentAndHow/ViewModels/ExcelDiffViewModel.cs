using System;
using System.Text;
using System.Windows.Input;
using WhatsDependentAndHow.Commands;
using Excel = Microsoft.Office.Interop.Excel;

namespace WhatsDependentAndHow.ViewModels
{
    public class ExcelDiffViewModel : PropertyChangedViewModel, IDisposable
    {
        private string _leftFilePath = string.Empty;
        public string LeftFilePath
        {
            get { return _leftFilePath; }
            set
            {
                _leftFilePath = value;
                OnPropertyChanged("LeftFilePath");
                UpdateIsExcelInfoLoaded();
            }
        }

        private void UpdateIsExcelInfoLoaded()
        {
            if (_leftFilePath != null && _leftFilePath.Trim().Length > 0 && _rightFilePath != null && _rightFilePath.Trim().Length > 0)
                IsExcelFileInfoLoaded = true;
        }

        private string _rightFilePath = string.Empty;
        public string RightFilePath
        {
            get { return _rightFilePath; }
            set
            {
                _rightFilePath = value;
                OnPropertyChanged("RightFilePath");
                UpdateIsExcelInfoLoaded();
            }
        }

        private bool _isExcelFileInfoLoaded = false;
        public bool IsExcelFileInfoLoaded
        {
            get { return _isExcelFileInfoLoaded; }
            set
            {
                _isExcelFileInfoLoaded = value;
                OnPropertyChanged("IsExcelFileInfoLoaded");
            }
        }

        private StringBuilder _diffSummary = new StringBuilder();
        public string DiffSummary
        {
            get { return _diffSummary.ToString(); }
            set
            {
                _diffSummary.Append(value);
                OnPropertyChanged("DiffSummary");
            }
        }

        private bool _isBusy = false;
        public bool IsBusy
        {
            get { return _isBusy; }
            set
            {
                _isBusy = value;
                OnPropertyChanged("IsBusy");
            }
        }

        private readonly MainViewModel _mainViewModel;
        private ButtonFileSelectorForComparisonClickCommand _buttonFileSelectorForComparisonClickCommand;
        private ButtonFindDifferencesClickCommand _buttonFindDifferencesClickCommand;

        public Excel.Application XlApp
        {
            get { return _mainViewModel.XlApp; }
            set
            {
                _mainViewModel.XlApp = value;
                OnPropertyChanged("XlApp");
            }
        }

        public Excel.Workbook XlLeftWorkBook
        {
            get { return _mainViewModel.XlLeftWorkBook; }
            set
            {
                _mainViewModel.XlLeftWorkBook = value;
                OnPropertyChanged("XlLeftWorkBook");
            }
        }

        public Excel.Workbook XlRightWorkBook
        {
            get { return _mainViewModel.XlRightWorkBook; }
            set
            {
                _mainViewModel.XlRightWorkBook = value;
                OnPropertyChanged("XlRightWorkBook");
            }
        }

        public ExcelDiffViewModel()
        {
            _buttonFileSelectorForComparisonClickCommand = new ButtonFileSelectorForComparisonClickCommand(this);
            _buttonFindDifferencesClickCommand = new ButtonFindDifferencesClickCommand(this);
        }

        public ExcelDiffViewModel(MainViewModel mainViewModel)
        {
            _mainViewModel = mainViewModel;
            _buttonFileSelectorForComparisonClickCommand = new ButtonFileSelectorForComparisonClickCommand(this);
            _buttonFindDifferencesClickCommand = new ButtonFindDifferencesClickCommand(this);
        }

        public ICommand ButtonFileSelectorForComparisonClickCommand
        {
            get { return _buttonFileSelectorForComparisonClickCommand; }
        }

        public ICommand ButtonFindDifferencesClickCommand
        {
            get { return _buttonFindDifferencesClickCommand; }
        }

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
