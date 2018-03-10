using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WhatsDependentAndHow
{
    public class CellsKeyedCollection : KeyedCollection<string, CellObject>
    {
        protected override string GetKeyForItem(CellObject item)
        {
            return item.Name;
        }
    }
}
