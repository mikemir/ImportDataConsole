using ImportDataConsole.ExcelHelper.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportDataConsole.ExcelHelper.Extensions
{
    public static class ExcelExtensions
    {
        public static string GetColumnAttrName(this object item, string name)
        {
            var result = item.GetType()
                .GetProperties()
                .Where(prop => prop.GetCustomAttributesData().Any(attr => attr.AttributeType == typeof(ColumnName)))
                .Select(prop => new { Prop = prop, Attr = prop.GetCustomAttributes(true).SingleOrDefault(attr => attr is ColumnName) as ColumnName })
                .SingleOrDefault(col => col.Attr.Name == name);

            return result?.Prop.Name;
        }
    }
}
