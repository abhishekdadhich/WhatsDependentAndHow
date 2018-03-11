using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace WhatsDependentAndHow
{
    public class ButtonFileSelectorCommand : ICommand
    {
        private MainWindowViewModel _mainWindowViewModel;

        public ButtonFileSelectorCommand(MainWindowViewModel mainWindowViewModel)
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
            _mainWindowViewModel.ClearControls();

            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Filter = "Excel worksheets|*.xls*";
            openFileDialog.InitialDirectory = @"C:\POC";
            openFileDialog.Title = "Select Excel File";

            if(openFileDialog.ShowDialog() == true)
            {
                _mainWindowViewModel.ExcelFileDetails = openFileDialog.FileName;
            }
        }
    }
}
