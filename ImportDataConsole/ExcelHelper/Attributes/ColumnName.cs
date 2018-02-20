using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportDataConsole.ExcelHelper.Attributes
{
    public class ColumnName : Attribute
    {
        public int Order { get; set; }
        public string Name { get; set; }
        public bool Border { get; set; }
        public string NumberFormat { get; set; }
        public string DateFormat { get; set; }

        public bool Flag { get; set; }
    }
}
