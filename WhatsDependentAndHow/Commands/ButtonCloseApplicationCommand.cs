using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using WhatsDependentAndHow.ViewModels;
using Excel = Microsoft.Office.Interop.Excel;

namespace WhatsDependentAndHow.Commands
{
    public class ButtonCloseApplicationCommand : ICommand
    {
        private TreeGeneratorViewModel _treeGeneratorViewModel;

        public ButtonCloseApplicationCommand(TreeGeneratorViewModel treeGeneratorViewModel)
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
            Excel.Application xlApp = _treeGeneratorViewModel.XlApp;
            Excel.Workbook xlWorkBook = _treeGeneratorViewModel.XlWorkBook;

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
