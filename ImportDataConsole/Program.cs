﻿using ImportDataConsole.Excel;
using ImportDataConsole.Excel.Attributes;
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
        [ImportDisplay("IDENTIFICADOR"), ColumnRequired]
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
        static void Main(string[] args)
        {
            var watch = new Stopwatch();
            var arrayBytes = File.ReadAllBytes("C:/ImportExcel/test.xlsx");

            watch.Start();
            var resultExcel = ExcelHelper.Import<Test>(arrayBytes);
            var resultErrorsExcel = resultExcel.Where(item => !item.IsValid).ToList();
            watch.Stop();

            Console.WriteLine($"Tiempo: {watch.ElapsedMilliseconds}");
        }
    }
}
