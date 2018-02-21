using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;

namespace ImportDataConsole.Excel.Attributes
{
    public abstract class ImportValidationAttribute
    {
        public string ColumnName { get; set; }
        public string ErrorMessage { get; set; }

        public bool IsValid(object value)
        {
            return true;
        }
    }
}
