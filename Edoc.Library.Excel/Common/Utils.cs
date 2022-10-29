using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Edoc.Library.Excel.Common
{
    internal static class Utils
    {
        public static T ConvertToOrDefault<T>(this object source, T def = default)
        {
            if (source == null)
                return def;

            if (source is T)
                return (T)source;

            return (T)Convert.ChangeType(source, Nullable.GetUnderlyingType(typeof(T)) ?? typeof(T));
        }

        public static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
    }
}
