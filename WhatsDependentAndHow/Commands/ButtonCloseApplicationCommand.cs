using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using Excel = Microsoft.Office.Interop.Excel;

namespace WhatsDependentAndHow.Commands
{
    public class ButtonCloseApplicationCommand : ICommand
    {
        private MainWindowViewModel _mainWindowViewModel;

        public ButtonCloseApplicationCommand(MainWindowViewModel mainWindowViewModel)
        {
            _mainWindowViewModel = mainWindowViewModel;
        }

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            Excel.Application xlApp = _mainWindowViewModel.XlApp;
            Excel.Workbook xlWorkBook = _mainWindowViewModel.XlWorkBook;

            if(xlWorkBook != null)
            {
                xlWorkBook.Close();
                Marshal.ReleaseComObject(xlWorkBook);
            }

            if(xlApp != null)
            {
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }

            System.Windows.Application.Current.Shutdown();
        }
    }
}
