using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using WhatsDependentAndHow.ViewModels;

namespace WhatsDependentAndHow.Commands
{
    public class ButtonFileSelectorForComparisonClickCommand : ICommand
    {
        private ExcelDiffViewModel _excelDiffViewModel;
        public ButtonFileSelectorForComparisonClickCommand(ExcelDiffViewModel excelDiffViewModel)
        {
            _excelDiffViewModel = excelDiffViewModel;
        }

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            OpenFileDialog openFileDialog;

            string side = parameter as string;

            if(side != null)
            {
                openFileDialog = Helpers.Helpers.GetExcelOpenFileDialog(string.Format("Select {0} Excel File", side));

                if (openFileDialog.ShowDialog() == true)
                {
                    if (side == "Left")
                        _excelDiffViewModel.LeftFilePath = openFileDialog.FileName;
                    else
                        _excelDiffViewModel.RightFilePath = openFileDialog.FileName;
                }
            }

            
        }
    }
}
