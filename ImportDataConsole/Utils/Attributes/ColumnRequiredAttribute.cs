using ClosedXML.Excel;
using ImportDataConsole.Excel.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportDataConsole.Utils.Attributes
{
    public class ColumnRequiredAttribute : ImportValidationAttribute
    {
        public override bool IsValid(IXLCell cell, string columName)
        {
            var valid = true;

            if (string.IsNullOrEmpty(cell.Value?.ToString()))
            {
                ErrorMessage = $"El campo \"{columName}\" es obligatorio (error en celda [{cell.Address.ColumnLetter}{cell.Address.RowNumber}])";
                valid = false;
            }

            return valid;
        }
    }
}
