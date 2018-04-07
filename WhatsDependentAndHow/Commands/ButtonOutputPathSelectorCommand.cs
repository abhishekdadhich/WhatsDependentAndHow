using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using WhatsDependentAndHow.ViewModels;

namespace WhatsDependentAndHow
{
    class ButtonOutputPathSelectorCommand : ICommand
    {
        private TreeGeneratorViewModel _treeGeneratorViewModel;

        public ButtonOutputPathSelectorCommand(TreeGeneratorViewModel treeGeneratorViewModel)
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
            using (var folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog())
            {
                folderBrowserDialog.ShowNewFolderButton = true;

                if(_treeGeneratorViewModel.ExcelFileDetails.Length > 0)
                {
                    folderBrowserDialog.SelectedPath = Path.GetDirectoryName(_treeGeneratorViewModel.ExcelFileDetails);
                    folderBrowserDialog.RootFolder = Environment.SpecialFolder.Desktop;
                }

                System.Windows.Forms.DialogResult result = folderBrowserDialog.ShowDialog();

                if(result == System.Windows.Forms.DialogResult.OK)
                {
                    _treeGeneratorViewModel.OutputFilePath = folderBrowserDialog.SelectedPath;
                }
            }
        }
    }
}
