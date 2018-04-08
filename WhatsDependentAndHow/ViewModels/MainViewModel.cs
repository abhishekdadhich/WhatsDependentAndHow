using MahApps.Metro.Controls;
using MahApps.Metro.IconPacks;
using Excel = Microsoft.Office.Interop.Excel;

namespace WhatsDependentAndHow.ViewModels
{
    public class MainViewModel : PropertyChangedViewModel
    {
        private HamburgerMenuItemCollection _menuItems;
        private HamburgerMenuItemCollection _menuOptionItems;

        public MainViewModel()
        {
            CreateMenuItems();
        }

        private Excel.Application _xlApp = null;
        public Excel.Application XlApp
        {
            get { return _xlApp; }
            set
            {
                _xlApp = value;
                OnPropertyChanged("XlApp");
            }
        }

        private Excel.Workbook _xlWorkBookForTreeGeneration = null;
        public Excel.Workbook XlWorkBookForTreeGeneration
        {
            get { return _xlWorkBookForTreeGeneration; }
            set
            {
                _xlWorkBookForTreeGeneration = value;
                OnPropertyChanged("XlWorkBookForTreeGeneration");
            }
        }

        private Excel.Workbook _xlLeftWorkBook = null;
        public Excel.Workbook XlLeftWorkBook
        {
            get { return _xlLeftWorkBook; }
            set
            {
                _xlLeftWorkBook = value;
                OnPropertyChanged("XlLeftWorkBook");
            }
        }

        private Excel.Workbook _xlRightWorkBook = null;
        public Excel.Workbook XlRightWorkBook
        {
            get { return _xlRightWorkBook; }
            set
            {
                _xlRightWorkBook = value;
                OnPropertyChanged("XlRightWorkBook");
            }
        }

        public void CreateMenuItems()
        {
            MenuItems = new HamburgerMenuItemCollection
            {
                new HamburgerMenuIconItem()
                {
                    Icon = new PackIconMaterial() {Kind = PackIconMaterialKind.FileTree},
                    Label = "Tree Generator",
                    ToolTip = "Generator Precedent or Dependents Tree",
                    Tag = new TreeGeneratorViewModel(this)
                },
                new HamburgerMenuIconItem()
                {
                    Icon = new PackIconMaterial() {Kind = PackIconMaterialKind.Compare},
                    Label = "Excel Diff",
                    ToolTip = "Spot difference in two Excel Workbooks",
                    Tag = new ExcelDiffViewModel(this)
                }
            };

            MenuOptionItems = new HamburgerMenuItemCollection
            {
                new HamburgerMenuIconItem()
                {
                    Icon = new PackIconMaterial() {Kind = PackIconMaterialKind.Power},
                    Label = "Exit",
                    ToolTip = "Exit Application",
                    Tag = new ExitApplicationViewModel(this)
                }
            };
        }

        public HamburgerMenuItemCollection MenuItems
        {
            get { return _menuItems; }
            set
            {
                if (Equals(value, _menuItems)) return;
                _menuItems = value;
                OnPropertyChanged();
            }
        }

        public HamburgerMenuItemCollection MenuOptionItems
        {
            get { return _menuOptionItems; }
            set
            {
                if (Equals(value, _menuOptionItems)) return;
                _menuOptionItems = value;
                OnPropertyChanged();
            }
        }
    }
}
