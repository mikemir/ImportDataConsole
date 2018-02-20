using ImportDataConsole.ExcelHelper.Attributes;
using ImportDataConsole.ExcelHelper.Extensions;
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
            var test = new Test();
            var result = test.GetColumnAttrName("IDENTIFICADOR");
            var result2 = test.GetColumnAttrName("VALOR");
            var result3 = test.GetColumnAttrName("test");
            var resultExcel = ExcelHelper.ExcelHelper.Import<Test>(arrayBytes);
        }
    }
}
