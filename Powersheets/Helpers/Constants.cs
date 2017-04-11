using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {

    internal static class Constants {

        internal static bool IsFormula(this string value) {
            if (!String.IsNullOrWhiteSpace(value)) {
                return (value[0] == '=') ? true : false;
            }
            return false;
        }

        internal static bool IsHeadingFormula(this string value) {
            if (value.IsFormula()) {
                string formulaType = value.ExtractFormulaType();
                return (formulaType == "HEADING") ? true : false;
            }
            return false;
        }

        internal static bool IsRowEmpty(this DataRow row, params int[] columnsToCheck) {
            int count = row.NonNullColumnCount(columnsToCheck);
            return (count == 0) ? true : false;
        }

        internal static T CastToType<T>(this object value) {
            try {
                T convertedValue = (T)Convert.ChangeType(value, typeof(T));
                return convertedValue;
            }
            catch {
                return default(T);
            }
        }

        internal static object CastToStringType(this object value, string type) {
            string valueString = (string)Convert.ChangeType(value, typeof(string));
            switch (type) {
                case "DECIMAL":
                    decimal de;
                    if (decimal.TryParse(valueString, out de)) {
                        return de;
                    }
                    return null;
                case "DOUBLE":
                    double du;
                    if (double.TryParse(valueString, out du)) {
                        return du;
                    }
                    return null;
                case "INT":
                    int it;
                    if (int.TryParse(valueString, out it)) {
                        return it;
                    }
                    return null;
                case "DATE":
                case "DATETIME":
                    DateTime dt;
                    if (DateTime.TryParse(valueString, out dt)) {
                        return dt;
                    }
                    else {
                        double dut;
                        if (double.TryParse(valueString, out dut)) {
                            return DateTime.FromOADate(dut);
                        }
                    }
                    return null;
                case "STRING":
                    return valueString;
                default:
                    return null;
            }
        }

        internal static object CastToPropertyType(this object value, PropertyInfo p) {
            if (value == null || value.GetType() == typeof(DBNull)) {
                return null;
            }

            string valueString = (string)Convert.ChangeType(value, typeof(string));
            if (p.PropertyType == typeof(decimal)) {
                decimal de;
                if (decimal.TryParse(valueString, out de)) {
                    return de;
                }
                return null;
            }
            else if (p.PropertyType == typeof(double)) {
                double du;
                if (double.TryParse(valueString, out du)) {
                    return du;
                }
                return null;
            }
            else if (p.PropertyType == typeof(int)) {
                int it;
                if (int.TryParse(valueString, out it)) {
                    return it;
                }
                return null;
            }
            else if (p.PropertyType == typeof(DateTime)) {
                DateTime dt;
                if (DateTime.TryParse(valueString, out dt)) {
                    return dt;
                }
                else {
                    // Handle that strange way of storing time in number of days since...
                    double dtu;
                    if (double.TryParse(valueString, out dtu)) {
                        return DateTime.FromOADate(dtu);
                    }
                }
                return null;
            }
            else if (p.PropertyType == typeof(string)) {
                return valueString;
            }
            else {
                return null;
            }
        }

        internal static string[] ColumnHeadings(this IEnumerable<IPowersheetPropertyMap> value, bool stripFormula = false) {
            var names = new List<string>();

            foreach (var v in value) {
                if (v.SpreadsheetMapValue.IsHeadingFormula() && stripFormula) {
                    names.Add(v.SpreadsheetMapValue.ExtractFormulaValue());
                }
                else if (!v.SpreadsheetMapValue.IsFormula()) {
                    names.Add(v.SpreadsheetMapValue);
                }
            }

            return names.ToArray();
        }

        internal static string ExtractFileExtension(this string value) {
            int start = value.NthIndexRightOf('.', 1) + 1;
            return new string(value.ToCharArray().Skip(start).ToArray());
        }

        internal static string ExtractFormulaType(this string value) {
            return new string(value.ToCharArray().Skip(1).TakeWhile(x => x != '(').ToArray()).ToUpper();
        }

        internal static string ExtractFormulaSubType(this string value) {
            return new string(value.ToCharArray().Skip(value.NthIndexOf('(', 1, 1)).TakeWhile(x => x != '(').ToArray()).ToUpper();
        }

        internal static string ExtractFormulaCastType(this string value) {
            int start = value.NthIndexRightOf(')', 2, -1);
            return new string(value.ToCharArray().Skip(start).TakeWhile(x => x != ')').ToArray()).ToUpper();
        }

        internal static string ExtractFormulaValue(this string value) {
            int start = value.NthIndexOf('(', 2, 1);
            int end = value.NthIndexRightOf(')', 2, start);
            return new string(value.ToCharArray().Skip(start).Take(end).ToArray());
        }

        internal static string ToStringValue(this object value) {
            if (value == null) {
                return string.Empty;
            }

            if (value.GetType() == typeof(string)) {
                return (string)value;
            }
            else if (value.GetType() == typeof(DateTime)) {
                return ((DateTime)value).ToShortDateString();
            }
            else {
                return value.ToString();
            }
        }

        internal static int[] ColumnIndexes(this IEnumerable<IPowersheetPropertyMap> value) {
            var indexes = new List<int>();

            foreach (var v in value) {
                indexes.Add(v.ColumnIndex);
            }

            return indexes.ToArray();
        }

        internal static int ColumnIndexOf(this DataRow row, string heading) {
            if (heading.IsFormula()) {
                if (heading.IsHeadingFormula()) {
                    string subType = heading.ExtractFormulaSubType();
                    string value = heading.ExtractFormulaValue();

                    value = value.ToLower().Trim();
                    for (int i = 0; i < row.ItemArray.Count(); i++) {
                        object obj = row.ItemArray[i];
                        try {
                            string objString = obj.ToString().ToLower().Trim();
                            switch (subType) {
                                case "LIKE":
                                    int computeLength;
                                    string cast = heading.ExtractFormulaCastType();
                                    if (int.TryParse(cast, out computeLength)) {
                                        int compute = value.LevenshteinCompute(objString);
                                        if (compute <= computeLength) {
                                            return i;
                                        }
                                    }
                                    break;
                                case "EQUALS":
                                    if (value == objString) {
                                        return i;
                                    }
                                    break;
                                default:
                                    continue;
                            }
                        }
                        catch {
                            continue;
                        }
                    }
                }
            }
            else {
                heading = heading.ToLower().Trim();
                for (int i = 0; i < row.ItemArray.Count(); i++) {
                    object obj = row.ItemArray[i];
                    try {
                        string objString = obj.ToString().ToLower().Trim();
                        if (heading == objString) {
                            return i;
                        }
                    }
                    catch {
                        continue;
                    }
                }
            }
            return -1;
        }

        internal static int TableIndexOf(this DataSet data, string tableName) {
            for (int i = 0; i < data.Tables.Count; i++) {
                if (tableName == data.Tables[i].ToString()) {
                    return i;
                }
            }
            return -1;
        }

        internal static int LevenshteinCompute(this string value, string compare) {
            if (string.IsNullOrEmpty(value)) {
                if (string.IsNullOrEmpty(compare)) {
                    return 0;
                }
                return compare.Length;
            }

            if (string.IsNullOrEmpty(compare)) {
                return value.Length;
            }

            int n = value.Length;
            int m = compare.Length;
            int[,] d = new int[n + 1, m + 1];

            for (int i = 0; i <= n; d[i, 0] = i++) ;
            for (int j = 1; j <= m; d[0, j] = j++) ;

            for (int i = 1; i <= n; i++) {
                for (int j = 1; j <= m; j++) {
                    int cost = (compare[j - 1] == value[i - 1]) ? 0 : 1;
                    int min1 = d[i - 1, j] + 1;
                    int min2 = d[i, j - 1] + 1;
                    int min3 = d[i - 1, j - 1] + cost;
                    d[i, j] = Math.Min(Math.Min(min1, min2), min3);
                }
            }
            return d[n, m];
        }

        internal static int MaxWidth(this DataTable table) {
            int width = -1;
            for (int i = 0; i < table.Rows.Count; i++) {
                if (table.Rows[i].ItemArray.Count() > width) {
                    width = table.Rows[i].ItemArray.Count();
                }
            }
            return width;
        }

        internal static int NonNullColumnCount(this DataRow row, params int[] columnsToInclude) {
            int count = 0;
            object[] rowData = row.ItemArray;

            if (columnsToInclude == null || columnsToInclude.Length == 0) {
                // Check the entire row
                for (int i = 0; i < rowData.Count(); i++) {
                    if (rowData[i] != null && rowData[i].GetType() != typeof(DBNull)) {
                        count++;
                    }
                }
            }
            else {
                // Check only the ones we want too!
                foreach (int i in columnsToInclude) {
                    if (i >= rowData.Count() || i < 0) {
                        continue;
                    }

                    if (rowData[i] != null && rowData[i].GetType() != typeof(DBNull)) {
                        count++;
                    }
                }
            }

            return count;
        }

        internal static int NthIndexOf(this string value, char t, int n, int p = 0) {
            return value.TakeWhile(c => (n -= (c == t ? 1 : 0)) > 0).Count() + p;
        }

        internal static int NthIndexRightOf(this string value, char t, int n, int p = 0) {
            int count = 0;
            for (int i = value.Length - 1; i >= 0; i--) {
                if (value[i] == t) {
                    count++;

                    if (count == n) {
                        return (i - p);
                    }
                }
            }
            return -1;
        }
    }
}
