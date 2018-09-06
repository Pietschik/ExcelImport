using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint;

namespace ExcelImport
{
    public class Column
    {
        public Column(string name, int index, Microsoft.SharePoint.SPFieldType type)
        {
            Name = name;
            Index = index;
            FieldType = type;
        }

        

        public  string Name { get; set; }
        public  int Index { get; set;}
        public  Microsoft.SharePoint.SPFieldType FieldType { get; set; }
    }
}
