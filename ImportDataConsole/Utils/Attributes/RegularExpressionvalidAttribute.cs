using ImportDataConsole.Excel.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace ImportDataConsole.Utils.Attributes
{
    public class RegularExpressionValidAttribute : ImportValidationAttribute
    {
        public override bool IsValid(IXLCell cell, string columName)
        {
            throw new NotImplementedException();
        }
    }
}
