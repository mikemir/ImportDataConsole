using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportDataConsole.Excel.Attributes
{
    public class ImportDisplayAttribute : Attribute
    {
        public ImportDisplayAttribute(string name)
        {
            ColumnName = name;
        }
        public string ColumnName { get; set; }
    }
}
