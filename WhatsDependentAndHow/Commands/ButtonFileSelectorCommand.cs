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

            try { Logger.SetLogWriter(new LogWriterFactory().Create()); } catch { }
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

            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = ConfigurationManager.AppSettings["OpenDialogFilter"].ToString();
            openFileDialog.InitialDirectory = @ConfigurationManager.AppSettings["DefaultDirectory"].ToString();
            openFileDialog.Title = "Select Excel File";

            Excel.Application xlApp = _treeGeneratorViewModel.XlApp;
            Excel.Workbook xlWorkBook = _treeGeneratorViewModel.XlWorkBook;

            if(openFileDialog.ShowDialog() == true)
            {
                _treeGeneratorViewModel.ExcelFileDetails = openFileDialog.FileName;

                try
                {
                    var swFileOpen = Stopwatch.StartNew();

                    OpenExcelFile(out xlApp, out xlWorkBook);

                    UpdateStatus(string.Format("File opened successfully in {0}. Total worksheets in file: {1}", swFileOpen.Elapsed, xlWorkBook.Worksheets.Count));

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

                    UpdateStatus(string.Format("Error: {0}.\n Stack Trace: {2}", message, e.Message, e.StackTrace));
                }
            }

            _treeGeneratorViewModel.IsBusy = false;
        }

        private void OpenExcelFile(out Excel.Application xlApp, out Excel.Workbook xlWorkBook)
        {
            try
            {
                xlApp = new Excel.Application();
                xlApp.DisplayAlerts = false;
                xlApp.AskToUpdateLinks = false;
                xlWorkBook = xlApp.Workbooks.Open(_treeGeneratorViewModel.ExcelFileDetails, UpdateLinks: false, ReadOnly: true);

                _treeGeneratorViewModel.XlApp = xlApp;
                _treeGeneratorViewModel.XlWorkBook = xlWorkBook;
            }
            catch (Exception e)
            {
                xlApp = null;
                xlWorkBook = null;
                UpdateStatus(string.Format("Error: {0}.\n Stack Trace: {1}", e.Message, e.StackTrace));
            }
        }

        private void UpdateStatus(string message)
        {
            Logger.Write(message);
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                _treeGeneratorViewModel.StatusMessage += string.Format("[{0} {1}]: {2}\n", DateTime.Now.ToShortDateString(), DateTime.Now.ToShortTimeString(), message);
            });
        }
    }
}
