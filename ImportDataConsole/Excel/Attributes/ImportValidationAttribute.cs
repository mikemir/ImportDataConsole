using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel.DataAnnotations;
using ClosedXML.Excel;

namespace ImportDataConsole.Excel.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public abstract class ImportValidationAttribute : Attribute
    {
        public string ErrorMessage { get; set; }

        public abstract bool IsValid(IXLCell cell, string columName);
    }
}
