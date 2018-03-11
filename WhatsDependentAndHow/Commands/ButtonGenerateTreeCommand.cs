using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace WhatsDependentAndHow.Commands
{
    public class ButtonGenerateTreeCommand : ICommand
    {
        private MainWindowViewModel _mainWindowViewModel;

        public ButtonGenerateTreeCommand(MainWindowViewModel mainWindowViewModel)
        {
            _mainWindowViewModel = mainWindowViewModel;
        }

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            if (_mainWindowViewModel.IsCellAddressValid)
                return true;
            else
                return false;
        }

        public void Execute(object parameter)
        {
            // TODO: Add Excel Parser Code
        }
    }
}
