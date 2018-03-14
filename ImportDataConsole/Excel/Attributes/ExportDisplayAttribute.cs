using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportDataConsole.Excel.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExportDisplayAttribute : Attribute
    {
        public ExportDisplayAttribute(string name)
        {
            if (name == null)
                throw new ArgumentNullException(nameof(name));

            Name = name;
        }

        public int Order { get; set; }
        public string Name { get; set; }
        public bool Border { get; set; }
        public string NumberFormat { get; set; }
        public string DateFormat { get; set; }
    }
}
