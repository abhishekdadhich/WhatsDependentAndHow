using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using WhatsDependentAndHow.ViewModels;
using Excel = Microsoft.Office.Interop.Excel;

namespace WhatsDependentAndHow.Commands
{
    public class ButtonFindDifferencesClickCommand : ICommand
    {
        private ExcelDiffViewModel _excelDiffViewModel;
        public ButtonFindDifferencesClickCommand(ExcelDiffViewModel excelDiffViewModel)
        {
            _excelDiffViewModel = excelDiffViewModel;
        }

        private Dictionary<string, int> _leftWorkBookSheets = new Dictionary<string, int>();
        private Dictionary<string, int> _rightWorkBookSheets = new Dictionary<string, int>();
        private Dictionary<string, int> _commonInBothSheets = new Dictionary<string, int>();
        private Dictionary<string, int> _onlyInLeftSheets = new Dictionary<string, int>();
        private Dictionary<string, int> _onlyInRightSheets = new Dictionary<string, int>();

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            if (!Validate())
                return;

            _excelDiffViewModel.IsBusy = true;

            if (!OpenExcelFiles())
            {
                _excelDiffViewModel.IsBusy = false;
                return;
            }

            PopulateSheets();
            PopulateDiffSummary();

            _excelDiffViewModel.IsBusy = false;
        }

        private void PopulateDiffSummary()
        {
            _excelDiffViewModel.DiffSummary = string.Format("Left Workbook Sheet Count: {0}\tRight Workbook Sheet Count:{1}\n", _leftWorkBookSheets.Count, _rightWorkBookSheets.Count);
            if (_rightWorkBookSheets.Count != _commonInBothSheets.Count)
            {
                string sheetsOnlyInLeft = string.Empty;
                foreach (string sheetName in _onlyInLeftSheets.Keys)
                {
                    sheetsOnlyInLeft += (sheetName + ",");
                }
                sheetsOnlyInLeft = sheetsOnlyInLeft.TrimEnd(',');

                string sheetsOnlyInRight = string.Empty;
                foreach (string sheetName in _onlyInRightSheets.Keys)
                {
                    sheetsOnlyInRight += (sheetName + ",");
                }
                sheetsOnlyInRight = sheetsOnlyInRight.TrimEnd(',');

                if (sheetsOnlyInLeft.Trim().Length > 0)
                    _excelDiffViewModel.DiffSummary = "Sheets missing in Right Workbook: " + sheetsOnlyInLeft;

                if (sheetsOnlyInRight.Trim().Length > 0)
                    _excelDiffViewModel.DiffSummary = "Sheets missing in Left Workbook: " + sheetsOnlyInRight;
            }
        }

        private void PopulateSheets()
        {
            foreach(Excel.Worksheet ws in _excelDiffViewModel.XlLeftWorkBook.Worksheets)
            {
                _leftWorkBookSheets.Add(ws.Name, ws.UsedRange.Count);
            }

            foreach (Excel.Worksheet ws in _excelDiffViewModel.XlRightWorkBook.Worksheets)
            {
                _rightWorkBookSheets.Add(ws.Name, ws.UsedRange.Count);
            }

            // find sheets that are present in both dictionaries
            _commonInBothSheets = _rightWorkBookSheets.Intersect(_leftWorkBookSheets).ToDictionary(x => x.Key, x => x.Value);

            // find sheets that are present only in Left but not in Right
            _onlyInLeftSheets = _leftWorkBookSheets.Except(_rightWorkBookSheets).ToDictionary(x => x.Key, x => x.Value);

            // find sheets that are present only in Right but not in Left
            _onlyInRightSheets = _rightWorkBookSheets.Except(_leftWorkBookSheets).ToDictionary(x => x.Key, x => x.Value);
        }

        private bool OpenExcelFiles()
        {
            try
            {
                Excel.Application xlApp;
                Excel.Workbook xlLeftWorkBook;
                Excel.Workbook xlRightWorkBook;

                Helpers.Helpers.OpenExcelFile(out xlApp, out xlLeftWorkBook, _excelDiffViewModel.LeftFilePath);

                if(xlApp != null)
                {
                    _excelDiffViewModel.XlApp = xlApp;
                    _excelDiffViewModel.UpdateStatus("Left File: XlApp initialized successfully");
                }
                else
                {
                    _excelDiffViewModel.UpdateStatus("Left File: XlApp is null. Can't proceed.");
                    return false;
                }

                if (xlLeftWorkBook != null)
                {
                    _excelDiffViewModel.XlLeftWorkBook = xlLeftWorkBook;
                    _excelDiffViewModel.UpdateStatus("Left Excel File Opened Successfully");
                }
                else
                {
                    _excelDiffViewModel.UpdateStatus("Left Excel File is null. Can't proceed.");
                    return false;
                }

                Helpers.Helpers.OpenExcelFile(out xlApp, out xlRightWorkBook, _excelDiffViewModel.RightFilePath);

                if (xlApp != null)
                {
                    _excelDiffViewModel.XlApp = xlApp;
                    _excelDiffViewModel.UpdateStatus("RightFile: XlApp initialized successfully");
                }
                else
                {
                    _excelDiffViewModel.UpdateStatus("RightFile: XlApp is null. Can't proceed.");
                    return false;
                }

                if (xlRightWorkBook != null)
                {
                    _excelDiffViewModel.XlRightWorkBook = xlRightWorkBook;
                    _excelDiffViewModel.UpdateStatus("Right Excel File Opened Successfully");
                }
                else
                {
                    _excelDiffViewModel.UpdateStatus("Right Excel File is null. Can't proceed.");
                    return false;
                }
            }
            catch (Exception e)
            {
                string message = "";

                if (e.InnerException != null)
                    message = e.InnerException.Message;

                _excelDiffViewModel.UpdateStatus(string.Format("Error: {0}.\n Stack Trace: {2}", message, e.Message, e.StackTrace));

                return false;
            }

            return true;
        }

        private bool Validate()
        {
            if (_excelDiffViewModel.LeftFilePath == null || _excelDiffViewModel.LeftFilePath.Trim().Length == 0)
            {
                _excelDiffViewModel.UpdateStatus("Left File not provided. Can't continue.");
                return false;
            }

            if (_excelDiffViewModel.RightFilePath == null || _excelDiffViewModel.RightFilePath.Trim().Length == 0)
            {
                _excelDiffViewModel.UpdateStatus("Right File not provided. Can't continue");
                return false;
            }

            if(_excelDiffViewModel.LeftFilePath == _excelDiffViewModel.RightFilePath)
            {
                _excelDiffViewModel.UpdateStatus("Both file paths are same. Won't continue.");
                return false;
            }

            return true;
        }
    }
}
