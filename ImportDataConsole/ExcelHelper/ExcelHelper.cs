﻿using ClosedXML.Excel;
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
        public static IEnumerable<TResult> Import<TResult>(byte[] data, int numRowHeader = 1) where TResult : new()
        {
            var result = new List<TResult>();
            var numRowData = numRowHeader + 1;

            using (var workBook = new XLWorkbook(new MemoryStream(data)))
            {
                var workSheet = workBook.Worksheets.FirstOrDefault();
                var rowHeader = workSheet.Row(numRowHeader);

                workSheet.Rows(rowHeader.RowNumber() + 1, workSheet.LastRowUsed().RowNumber())
                .ForEach(row => {
                    var item = new TResult();

                    row.Cells(1, row.LastCellUsed().Address.ColumnNumber)
                    .ForEach(cell => {
                        var cellHeader = workSheet.Cell(numRowHeader, cell.Address.ColumnNumber);
                        var propName = item.GetColumnAttrName(cellHeader.Value.ToString());

                        if (!cell.IsEmpty() && propName != null)
                        {
                            var prop = typeof(TResult).GetProperty(propName);
                            prop?.SetValue(item, Convert.ChangeType(cell.Value, prop.PropertyType));
                        }

                    });

                    result.Add(item);
                    numRowData++;
                });
            }

            return result;
        }
    }
}
