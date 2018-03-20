using ClosedXML.Excel;
using ImportDataConsole.Excel.Attributes;
using ImportDataConsole.Excel.Entities;
using ImportDataConsole.Excel.Exceptions;
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
    public static class ExcelWrapper
    {
        public static bool ValidateExcel(byte[] data, IEnumerable<string> columns, out string message)
        {
            bool result = true;
            message = string.Empty;

            var workbook = new XLWorkbook(new MemoryStream(data));
            var worksheet = workbook.Worksheets.FirstOrDefault();

            if (worksheet.IsEmpty())
            {
                message = "Excel esta vacío";
                return false;
            }

            var firstCell = worksheet.FirstCellUsed();
            var lastColumn = worksheet.LastColumnUsed();
            var range = worksheet.Range(firstCell.Address, lastColumn.LastCellUsed().Address);

            foreach (var column in columns)
            {
                var cellColumn = range.Search(column, CompareOptions.IgnoreCase, false).FirstOrDefault();

                if (cellColumn != null && column.Equals(cellColumn.Value.ToString()))
                {
                    var cellData = worksheet.Cell(cellColumn.Address.RowNumber + 1, cellColumn.Address.ColumnNumber);
                    if (cellData.IsEmpty())
                    {
                        result = false;
                        message = $"El archivo no tiene la lista de {cellColumn.Value} a generar";
                    }
                }
                else
                {
                    result = false;
                    message = $"El archivo no tiene la columna con encabezado  \"{column}\"";
                }

            }

            return result;
        }

        #region EXPORT
        public static byte[] Export<T>(IEnumerable<ExportExcel<T>> exportData) where T : class, new()
        {
            var ms = new MemoryStream();

            using (var workbook = new XLWorkbook())
            {
                exportData.ForEach(data => {
                    var worksheet = workbook.AddWorksheet(data.Detaills.Any() ? data.WorkSheet : $"{data.WorkSheet} (EMPTY)");
                    var lastRowHeader = worksheet.DrawHeaderOrFooter<int>(0); //ToDo: Crear Encabezao
                    var lastRowDataTable = worksheet.DrawDataTable(data.Detaills, lastRowHeader);
                    var lastRowFooter = worksheet.DrawHeaderOrFooter<int>(0, lastRowDataTable);//ToDo: Crear Pie
                });

                workbook.SaveAs(ms);
            }

            return ms.ToArray();
        }

        public static byte[] ExportWithTemplate(IEnumerable<ExportTemplateExcel> exportData, byte[] excelTemplate)
        {
            var ms = new MemoryStream();

            using (var workbook = new XLWorkbook(new MemoryStream(excelTemplate)))
            {
                workbook.Worksheets.ForEach(worksheet => {
                    var data = exportData.SingleOrDefault(export => export.WorkSheet.Equals(worksheet.Name));
                    if (data == null)
                        throw new NotFoundWorksheetExportException(worksheet.Name);

                    data.Detaills.ForEach(item => {
                        var startCell = worksheet.CellsUsed().SingleOrDefault(cell => cell.Value.ToString().Equals("{#" + item.SearchKey + "#}"));
                        if (item.IsTable && startCell != null)
                        {
                            var address = startCell.Address;
                            worksheet.DrawDataTable(item.GetDataTable(), address.RowNumber, address.ColumnNumber);
                        }
                        else if (startCell != null)
                        {
                            startCell.Value = item.Value;
                        }
                    });
                });

                workbook.SaveAs(ms);
            }

            return ms.ToArray();
        }

        private static int DrawHeaderOrFooter<T>(this IXLWorksheet worksheet, object data, int rowNumberStart = 1, int colNumberStart = 1)
        {
            return 1;
        }

        private static int DrawDataTable<T>(this IXLWorksheet worksheet, IEnumerable<T> data, int rowNumberStart = 1, int colNumberStart = 1) where T : class, new()
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

            return rowNumber;
        }

        private static Dictionary<string, PropertyInfo> GetColumnList(Type genericType, params string[] visibleColummns)
        {
            if (genericType == null)
                throw new ArgumentNullException(nameof(genericType));

            var columnList = genericType.GetProperties()
                .Where(prop => prop.GetCustomAttribute<ExportDisplayAttribute>() != null && visibleColummns == null ||
                               prop.GetCustomAttribute<ExportDisplayAttribute>() != null && visibleColummns.Contains(prop.Name))
                .Select(prop =>
                    new
                    {
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

                var rows = workSheet.Rows(rowHeader.RowNumber() + 1, workSheet.LastRowUsed().RowNumber());
                Parallel.ForEach(rows, row => {
                    var itemImport = new ImportExcel<T>();

                    row.Cells(1, rowHeader.LastCellUsed().Address.ColumnNumber)
                    .ForEach(cell => {

                        var cellHeader = workSheet.Cell(numRowHeader, cell.Address.ColumnNumber);
                        itemImport = ValidateCell(cellHeader.Value.ToString(), cell, itemImport);

                    });

                    result.Add(itemImport);
                });
            }

            return result;
        }
        #endregion
    }
}
