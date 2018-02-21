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
        public static T GetAttribute<T>(this MemberInfo obj) where T : class
        {
            return obj.GetCustomAttributes(true).FirstOrDefault(param => param is T) as T;
        }

        public static IEnumerable<T> GetAttributes<T>(this MemberInfo obj) where T : class
        {
            return obj.GetCustomAttributes(true).Where(param => param is T).Select(item => item as T).ToList();
        }

        public static MemberInfo GetPropertyInfo<T>(this T obj, Expression<Func<T, object>> expression)
        {
            var uniExpresion = expression.Body as UnaryExpression;
            var memberExpresion = uniExpresion?.Operand as MemberExpression;

            return memberExpresion != null ? obj.GetType().GetProperty(memberExpresion.Member.Name) : null;
        }

        public static string GetPropertyNameByColumnAttr(this object item, string columnName)
        {
            var result = item.GetType()
                .GetProperties()
                .Where(prop => prop.GetCustomAttributesData().Any(attr => attr.AttributeType == typeof(ImportDisplayAttribute)))
                .Select(prop => new { Prop = prop, Attr = prop.GetCustomAttributes(true).SingleOrDefault(attr => attr is ImportDisplayAttribute) as ImportDisplayAttribute })
                .SingleOrDefault(col => col.Attr.ColumnName == columnName);

            return result?.Prop.Name;
        }
    }
}
