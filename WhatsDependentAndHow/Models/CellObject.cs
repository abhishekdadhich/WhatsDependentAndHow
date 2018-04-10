using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WhatsDependentAndHow
{
    public class CellObject
    {
        private string _name = string.Empty;
        public string Name
        {
            get { return _name; }
            set
            {
                if (value != null)
                    _name = value;
            }
        }

        private string _value = string.Empty;
        public string Value
        {
            get
            {
                return _value;
            }
            set
            {
                if (value != null)
                    _value = value;
            }
        }

        private string _formula = string.Empty;
        public string Formula
        {
            get { return _formula; }
            set
            {
                if (value != null)
                    _formula = value;
            }
        }

        private string _rowHeading = string.Empty;
        public string RowHeading
        {
            get { return _rowHeading; }
            set
            {
                if (value != null)
                    _rowHeading = value;
            }
        }

        public CellObject Parent { get; set; }

        private ObservableCollection<CellObject> _children = new ObservableCollection<CellObject>();
        public ObservableCollection<CellObject> Children
        {
            get { return _children; }
            set { _children = value; }
        }
    }
}
