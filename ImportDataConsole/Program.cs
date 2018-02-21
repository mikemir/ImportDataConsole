﻿using ImportDataConsole.Excel;
using ImportDataConsole.Excel.Attributes;
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

            var resultExcel = ExcelHelper.Import<Test>(arrayBytes);
        }
    }
}
