using ClosedXML.Excel;
using ImportDataConsole.ExcelHelper.Entities;
using ImportDataConsole.ExcelHelper.Extensions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportDataConsole.ExcelHelper
{
    public static class ExcelHelper
    {
        public static byte[] Export<TData>(IEnumerable<ExportExcel<TData>> data) where TData : class, new()
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

        public static IEnumerable<ImportExcel<TResult>> Import<TResult>(byte[] data, int numRowHeader = 1) where TResult : class, new()
        {
            var numRowData = numRowHeader + 1;
            var result = new List<ImportExcel<TResult>>();

            using (var workBook = new XLWorkbook(new MemoryStream(data)))
            {
                var workSheet = workBook.Worksheets.FirstOrDefault();
                var rowHeader = workSheet.Row(numRowHeader);

                workSheet.Rows(rowHeader.RowNumber() + 1, workSheet.LastRowUsed().RowNumber())
                .ForEach(row => {
                    var itemImport = new ImportExcel<TResult>();

                    row.Cells(1, row.LastCellUsed().Address.ColumnNumber)
                    .ForEach(cell => {
                        var cellHeader = workSheet.Cell(numRowHeader, cell.Address.ColumnNumber);
                        var propName = itemImport.Item.GetColumnAttrName(cellHeader.Value.ToString());

                        if (!cell.IsEmpty() && propName != null)
                        {
                            var prop = typeof(TResult).GetProperty(propName);
                            prop?.SetValue(itemImport.Item, Convert.ChangeType(cell.Value, prop.PropertyType));
                        }

                    });

                    result.Add(itemImport);
                    numRowData++;
                });
            }

            return result;
        }
    }
}
