using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WhatsDependentAndHow
{
    public class CellObject
    {
        public string Name { get; set; }
        public string Value { get; set; }
        public string Formula { get; set; }

        public List<CellObject> Dependents = new List<CellObject>();
        public List<CellObject> Precedents = new List<CellObject>();
    }
}
