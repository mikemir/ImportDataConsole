using ImportDataConsole.Excel.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ImportDataConsole.Excel.Extensions
{
    public static class ExcelExtensions
    {
        public static MemberInfo GetPropertyInfo<T>(this T obj, Expression<Func<T, object>> expression)
        {
            var uniExpresion = expression.Body as UnaryExpression;
            var memberExpresion = uniExpresion?.Operand as MemberExpression;

            return memberExpresion == null ? null : obj.GetType().GetProperty(memberExpresion.Member.Name);
        }

        public static string GetPropertyNameByColumnAttr(this object obj, string columnName)
        {
            var result = obj.GetType()
                .GetProperties()
                .Where(prop => prop.GetCustomAttributesData().Any(attr => attr.AttributeType == typeof(ImportDisplayAttribute)))
                .Select(prop => new { Prop = prop, Attr = prop.GetCustomAttribute<ImportDisplayAttribute>() })
                .FirstOrDefault(col => col.Attr.ColumnName == columnName);

            return result?.Prop.Name;
        }
    }
}
