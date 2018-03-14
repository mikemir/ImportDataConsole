using ImportDataConsole.Excel.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.Text.RegularExpressions;

namespace ImportDataConsole.Utils.Attributes
{
    public class RegularExpressionValidAttribute : ImportValidationAttribute
    {
        public string Pattern { get; set; }
        public RegularExpressionValidAttribute(string pattern)
        {
            Pattern = pattern;
        }
        public override bool IsValid(IXLCell cell, string columName)
        {
            var result = true;

            if (!string.IsNullOrEmpty(Pattern))
            {
                var match = Regex.Match(cell.Value.ToString(), Pattern);

                if (!match.Success)
                {
                    ErrorMessage = $"El valor \"{cell.Value}\" no es un valor válido (error en celda [{cell.Address.ColumnLetter}{cell.Address.RowNumber}])";
                    result = match.Success;
                }
            }

            return result;
        }
    }
}
