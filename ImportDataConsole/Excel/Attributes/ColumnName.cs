using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportDataConsole.Excel.Attributes
{
    public class ColumnNameAttribute : Attribute
    {
        public int Order { get; set; }
        public string Name { get; set; }
        public bool Border { get; set; }
        public string NumberFormat { get; set; }
        public string DateFormat { get; set; }
    }
}
