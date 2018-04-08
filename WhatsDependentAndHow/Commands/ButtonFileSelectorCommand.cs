using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using Microsoft.Practices.EnterpriseLibrary.Logging;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using WhatsDependentAndHow.ViewModels;

namespace WhatsDependentAndHow
{
    public class ButtonFileSelectorCommand : ICommand
    {
        private TreeGeneratorViewModel _treeGeneratorViewModel;

        public ButtonFileSelectorCommand(TreeGeneratorViewModel treeGeneratorViewModel)
        {
            _treeGeneratorViewModel = treeGeneratorViewModel;
        }

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            _treeGeneratorViewModel.ClearControls();
            _treeGeneratorViewModel.IsBusy = true;

            Excel.Application xlApp = _treeGeneratorViewModel.XlApp;
            Excel.Workbook xlWorkBook = _treeGeneratorViewModel.XlWorkBook;

            OpenFileDialog openFileDialog = Helpers.Helpers.GetExcelOpenFileDialog("Select Excel File");

            if(openFileDialog.ShowDialog() == true)
            {
                _treeGeneratorViewModel.ExcelFileDetails = openFileDialog.FileName;

                try
                {
                    var swFileOpen = Stopwatch.StartNew();

                    Helpers.Helpers.OpenExcelFile(out xlApp, out xlWorkBook, openFileDialog.FileName);
                    _treeGeneratorViewModel.XlApp = xlApp;
                    _treeGeneratorViewModel.XlWorkBook = xlWorkBook;

                    _treeGeneratorViewModel.UpdateStatus(string.Format("File opened successfully in {0}. Total worksheets in file: {1}", swFileOpen.Elapsed, xlWorkBook.Worksheets.Count));

                    foreach(Excel.Worksheet sheet in xlWorkBook.Worksheets)
                    {
                        _treeGeneratorViewModel.WorkSheetNames.Add(sheet.Name);
                    }
                }
                catch(Exception e)
                {
                    string message = "";

                    if (e.InnerException != null)
                        message = e.InnerException.Message;

                    _treeGeneratorViewModel.UpdateStatus(string.Format("Error: {0}.\n Stack Trace: {2}", message, e.Message, e.StackTrace));
                }
            }

            _treeGeneratorViewModel.IsBusy = false;
        }
    }
}
