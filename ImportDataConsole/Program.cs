using ImportDataConsole.Excel;
using ImportDataConsole.Excel.Attributes;
using ImportDataConsole.Excel.Entities;
using ImportDataConsole.Excel.Extensions;
using ImportDataConsole.Utils.Attributes;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportDataConsole
{
    public class Test
    {
        [ExportDisplay("IDENTIFICADOR")]
        [ImportDisplay("IDENTIFICADOR")]
        public int Id { get; set; }

        [ExportDisplay("VALOR")]
        [ImportDisplay("VALOR"), ColumnRequired]
        public string Nombre { get; set; }

        [ExportDisplay("FECHA")]
        [ImportDisplay("FECHA"), DateValid, ColumnRequired]
        public DateTime Fecha { get; set; }
    }

    class Program
    {
        public static List<Test> GenerateData()
        {
            var result = new List<Test>();

            for (int i = 0; i < 100; i++)
            {
                result.Add(new Test
                {
                    Id = i,
                    Nombre = Guid.NewGuid().ToString("N"),
                    Fecha = DateTime.Now.AddDays(-10).AddDays(i)
                });
            }

            return result;
        }

        static void Main(string[] args)
        {
            var watch = new Stopwatch();
            var excelPath = "C:/Excel/test.xlsx";

            watch.Start();
            var bytes = ExcelHelper.Export(new ExportExcel<Test>[] { new ExportExcel<Test>("DATA IMPORT", GenerateData()) });
            File.WriteAllBytes(excelPath, bytes);
            watch.Stop();
            Console.WriteLine($"Tiempo: {watch.Elapsed}");
            watch.Reset();

            watch.Start();
            //var arrayBytes = File.ReadAllBytes(excelPath);
            var arrayBytes = File.ReadAllBytes("C:/Excel/test2.xlsx");
            var resultExcel = ExcelHelper.Import<Test>(arrayBytes);
            watch.Stop();

            var resultValidExcel = resultExcel.Where(item => item.IsValid).ToList();
            var resultNotValidsExcel = resultExcel.Where(item => !item.IsValid).ToList();

            Console.WriteLine($"Tiempo: {watch.Elapsed}");
            Console.Read();
        }
    }
}
