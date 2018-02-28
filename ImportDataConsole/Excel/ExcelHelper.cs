using ClosedXML.Excel;
using ImportDataConsole.Excel.Attributes;
using ImportDataConsole.Excel.Entities;
using ImportDataConsole.Excel.Extensions;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ImportDataConsole.Excel
{
    public static class ExcelHelper
    {
        #region EXPORT
        public static byte[] Export<T>(IEnumerable<ExportExcel<T>> exportData) where T : class, new()
        {
            var ms = new MemoryStream();

            using (var workbook = new XLWorkbook())
            {
                foreach (var data in exportData)
                {
                    var worksheet = workbook.AddWorksheet(data.Detaills.Any() ? data.WorkSheet : $"{data.WorkSheet} (EMPTY)");
                    //ToDo: Crear Encabezao
                    worksheet.AddDetails(data.Detaills);
                    //ToDo: Crear Pie
                }

                workbook.SaveAs(ms);
            }

            return ms.ToArray();
        }

        private static IXLWorksheet AddHeaderOrFooter<T>(this IXLWorksheet worksheet, object data, int colNumberStart = 1, int rowNumberStart = 1)
        {
            return worksheet;
        }

        private static IXLWorksheet AddDetails<T>(this IXLWorksheet worksheet, IEnumerable<T> data, int colNumberStart = 1, int rowNumberStart = 1) where T : class, new()
        {
            var colNumber = colNumberStart;
            var rowNumber = rowNumberStart;

            var first = data.FirstOrDefault();
            var dataProps = GetColumnList(first?.GetType(), null);

            dataProps.ForEach(p =>
            {
                var cell = worksheet.Cell(rowNumber, colNumber++);
                DrawHeaderCell(cell, p.Key);
            });

            data.ForEach(item =>
            {
                rowNumber++;
                colNumber = 1;
                dataProps.ForEach(p =>
                {
                    var cell = worksheet.Cell(rowNumber, colNumber++);
                    DrawDataCell(cell, p.Value.GetValue(item), p.Value.GetCustomAttribute<ExportDisplayAttribute>());
                });
            });

            worksheet.Columns(colNumberStart, dataProps.Count).AdjustToContents();

            return worksheet;
        }

        private static Dictionary<string, PropertyInfo> GetColumnList(Type genericType, params string[] visibleColummns)
        {
            if (genericType == null)
                throw new ArgumentNullException(nameof(genericType));

            var columnList = genericType.GetProperties()
                .Where(prop => prop.GetCustomAttribute<ExportDisplayAttribute>() != null && visibleColummns == null ||
                               prop.GetCustomAttribute<ExportDisplayAttribute>() != null && visibleColummns.Contains(prop.Name))
                .Select(prop => new {
                        Attribute = prop.GetCustomAttribute<ExportDisplayAttribute>(),
                        PropertyInfo = prop
                    })
                    .OrderBy(prop => prop.Attribute.Order)
                    .ToDictionary(item => item.Attribute.Name, item => item.PropertyInfo);

            return columnList;
        }

        private static void DrawHeaderCell(IXLCell cell, object value)
        {
            cell.Style.Font.Bold = true;
            cell.Style.Font.FontSize = 9;
            cell.Style.Font.FontName = "Arial";
            cell.Style.Fill.BackgroundColor = XLColor.Gainsboro;
            cell.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
            cell.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            cell.SetValue(value);
        }

        private static void DrawDataCell(IXLCell cell, object value, ExportDisplayAttribute attribute)
        {
            cell.Style.Border.OutsideBorder = attribute.Border ? XLBorderStyleValues.Thin : XLBorderStyleValues.None;
            if (attribute.NumberFormat != null) cell.Style.NumberFormat.SetFormat(attribute.NumberFormat);
            if (attribute.DateFormat != null) cell.Style.DateFormat.SetFormat(attribute.DateFormat);
            //if (attribute.Flag) cell.Style.Fill.SetBackgroundColor(XLColor.Yellow);
            cell.Style.Font.FontName = "Arial";
            cell.Style.Font.FontSize = 9;

            cell.Value = value;
        }
        #endregion

        #region IMPORT
        private static ImportExcel<T> ValidateCell<T>(string cellHeader, IXLCell cell, ImportExcel<T> importContent) where T : class, new()
        {
            var cellValid = true;
            var propName = importContent.Item.GetPropertyNameByColumnAttr(cellHeader);

            if (propName != null)
            {
                var prop = typeof(T).GetProperty(propName);
                var validations = prop.GetCustomAttributes<ImportValidationAttribute>();

                validations.ForEach(val => {
                    if (!val.IsValid(cell, cellHeader))
                    {
                        importContent.ValidationMessage = importContent.ValidationMessage == null ? val.ErrorMessage
                                                            : $"{importContent.ValidationMessage}, {val.ErrorMessage}";
                        cellValid = false;
                    }
                });

                importContent.IsValid = importContent.IsValid && cellValid;

                if (cellValid) prop?.SetValue(importContent.Item, Convert.ChangeType(cell.Value, prop.PropertyType));
            }

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

                    row.Cells(1, rowHeader.LastCellUsed().Address.ColumnNumber)
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
        #endregion
    }
}
