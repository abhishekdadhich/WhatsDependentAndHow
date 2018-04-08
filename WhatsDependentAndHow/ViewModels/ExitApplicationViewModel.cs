using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using WhatsDependentAndHow.Commands;

namespace WhatsDependentAndHow.ViewModels
{
    public class ExitApplicationViewModel : PropertyChangedViewModel
    {
        private readonly MainViewModel _mainViewModel;
        public MainViewModel ViewModelMainWindow
        {
            get { return _mainViewModel; }
        }

        private ButtonCloseApplicationCommand _buttonCloseApplicationCommand;
        public ICommand ButtonCloseApplicationCommand
        {
            get { return _buttonCloseApplicationCommand; }
        }

        public ExitApplicationViewModel(MainViewModel mainViewModel)
        {
            _mainViewModel = mainViewModel;
            _buttonCloseApplicationCommand = new ButtonCloseApplicationCommand(this);
        }

        public ExitApplicationViewModel()
        {
            _buttonCloseApplicationCommand = new ButtonCloseApplicationCommand(this);
        }
    }
}
