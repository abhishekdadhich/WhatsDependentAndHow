using System;
using System.Runtime.InteropServices;
using System.Windows.Input;
using WhatsDependentAndHow.ViewModels;
using Excel = Microsoft.Office.Interop.Excel;

namespace WhatsDependentAndHow.Commands
{
    public class ButtonCloseApplicationCommand : ICommand
    {
        private ExitApplicationViewModel _viewModel;

        public ButtonCloseApplicationCommand(ExitApplicationViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            Excel.Application xlApp = _viewModel.ViewModelMainWindow.XlApp;
            Excel.Workbook xlWorkBookTreeGenerator = _viewModel.ViewModelMainWindow.XlWorkBookForTreeGeneration;
            Excel.Workbook xlLeftWorkBook = _viewModel.ViewModelMainWindow.XlLeftWorkBook;
            Excel.Workbook xlRightWorkBook = _viewModel.ViewModelMainWindow.XlRightWorkBook;

            try
            {
                if (xlWorkBookTreeGenerator != null)
                {
                    xlWorkBookTreeGenerator.Close();
                    Marshal.ReleaseComObject(xlWorkBookTreeGenerator);
                }
            }
            catch (Exception) { }


            try
            {
                if (xlLeftWorkBook != null)
                {
                    xlLeftWorkBook.Close();
                    Marshal.ReleaseComObject(xlLeftWorkBook);
                }
            }
            catch (Exception) { }
            
            try
            {
                if (xlRightWorkBook != null)
                {
                    xlRightWorkBook.Close();
                    Marshal.ReleaseComObject(xlRightWorkBook);
                }
            }
            catch (Exception) { }

            try
            {
                if (xlApp != null)
                {
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);
                }
            }
            catch (Exception) { }

            System.Windows.Application.Current.Shutdown();
        }
    }
}
