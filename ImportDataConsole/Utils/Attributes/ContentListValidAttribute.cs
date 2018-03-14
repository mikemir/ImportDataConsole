using ClosedXML.Excel;
using ImportDataConsole.Excel.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportDataConsole.Utils.Attributes
{
    public class ContentListValidAttribute : ImportValidationAttribute
    {
        public IEnumerable<string> List { get; set; }

        public ContentListValidAttribute(params string[] list)
        {
            List = list.Select(item => item.ToUpper());
        }

        public override bool IsValid(IXLCell cell, string columName)
        {
            var result = true;
            if (!List.Contains(cell.Value?.ToString().ToUpper()))
            {
                ErrorMessage = $"El valor de columna {columName} no válido en celda [{cell.Address.ColumnLetter}{cell.Address.RowNumber}]";
                result = false;
            }
            return result;
        }
    }
}
