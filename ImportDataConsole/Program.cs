using ImportDataConsole.Excel;
using ImportDataConsole.Excel.Attributes;
using ImportDataConsole.Excel.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportDataConsole
{
    public class Test
    {
        [ColumnName(Name = "IDENTIFICADOR")]
        public int Id { get; set; }
        [ColumnName(Name = "VALOR")]
        public string Nombre { get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            var arrayBytes = File.ReadAllBytes("C:/ImportExcel/test.xlsx");
            var test = new Test { Id = 1, Nombre = "Michael " };
            var prop = test.GetPropertyInfo(item => item.Id).GetAttribute<ColumnName>();

            var resultExcel = ExcelHelper.Import<Test>(arrayBytes);
        }
    }
}
