using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportDataConsole.Excel.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ImportDisplayAttribute : Attribute
    {
        public ImportDisplayAttribute(string name)
        {
            if (name == null)
                throw new ArgumentNullException(nameof(name));

            ColumnName = name;
        }
        public string ColumnName { get; set; }
    }
}
