using ClosedXML.Excel;
using ImportDataConsole.Excel.Entities;
using ImportDataConsole.Excel.Extensions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportDataConsole.Excel
{
    public static class ExcelHelper
    {
        public static byte[] Export<T>(IEnumerable<ExportExcel<T>> data) where T : class, new()
        {
            var ms = new MemoryStream();

            using (var workbook = new XLWorkbook())
            {
                foreach (var hoja in data)
                {

                }

                workbook.SaveAs(ms);
            }

            return ms.ToArray();
        }

        private static ImportExcel<T> ValidateCell<T>(string cellHeader, IXLCell cell, ImportExcel<T> importContent) where T : class, new()
        {
            var valid = true;
            var propName = importContent.Item.GetPropertyNameByColumnAttr(cellHeader);

            if (propName != null)
            {
                if (!cell.IsEmpty())
                {
                    var prop = typeof(T).GetProperty(propName);
                    prop?.SetValue(importContent.Item, Convert.ChangeType(cell.Value, prop.PropertyType));
                }
                else
                {
                    importContent.ValidationMessage = $"Error[{cell.Address.ColumnLetter}:{cell.Address.RowNumber}]: La columna {cellHeader} está vacia.";
                    valid = false;
                }
            }

            importContent.IsValid = importContent.IsValid && valid;

            return importContent;
        }

        public static IEnumerable<ImportExcel<T>> Import<T>(byte[] data, int numRowHeader = 1) where T : class, new()
        {
            var numRowData = numRowHeader + 1;
            var result = new List<ImportExcel<T>>();

            using (var workBook = new XLWorkbook(new MemoryStream(data)))
            {
                var workSheet = workBook.Worksheets.FirstOrDefault();
                var rowHeader = workSheet.Row(numRowHeader);

                workSheet.Rows(rowHeader.RowNumber() + 1, workSheet.LastRowUsed().RowNumber())
                .ForEach(row => {
                    var itemImport = new ImportExcel<T>();

                    row.Cells(1, row.LastCellUsed().Address.ColumnNumber)
                    .ForEach(cell => {

                        var cellHeader = workSheet.Cell(numRowHeader, cell.Address.ColumnNumber);
                        itemImport = ValidateCell(cellHeader.Value.ToString(), cell, itemImport);

                    });

                    result.Add(itemImport);
                    numRowData++;
                });
            }

            return result;
        }
    }
}
