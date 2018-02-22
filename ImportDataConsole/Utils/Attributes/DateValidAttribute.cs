using ImportDataConsole.Excel.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace ImportDataConsole.Utils.Attributes
{
    public class DateValidAttribute : ImportValidationAttribute
    {
        public override bool IsValid(IXLCell cell, string columName)
        {
            DateTime fecha;
            var result = true;

            if(!DateTime.TryParse(cell.Value.ToString(), out fecha))
            {
                ErrorMessage = $"El valor \"{cell.Value}\" no es una fecha valida (error en celda [{cell.Address.ColumnLetter}{cell.Address.RowNumber}])";
                result = false;
            }

            return result;
        }
    }
}
