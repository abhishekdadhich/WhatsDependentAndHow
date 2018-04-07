using MahApps.Metro.Controls;
using MahApps.Metro.IconPacks;

namespace WhatsDependentAndHow.ViewModels
{
    public class MainViewModel : PropertyChangedViewModel
    {
        private HamburgerMenuItemCollection _menuItems;
        private HamburgerMenuItemCollection _menuOptionItems;

        public MainViewModel()
        {
            this.CreateMenuItems();
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
                    Tag = new TreeGeneratorViewModel(this)
                }
            };

            MenuOptionItems = new HamburgerMenuItemCollection
            {
                new HamburgerMenuIconItem()
                {
                    Icon = new PackIconMaterial() {Kind = PackIconMaterialKind.Power},
                    Label = "Exit",
                    ToolTip = "Exit Application",
                    Tag = new TreeGeneratorViewModel(this)
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
