using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {

    public sealed class SpreadsheetImporter<T> : IPowersheetImporter<T> {

        private DataSet _data;

        public DataSet Data { get { return _data; } set { throw new Exception("Data is a read-only field!"); } }
        public int RowConfirmPercentage { get; set; }
        public int ValidateConfirmPercentage { get; set; }

        public SpreadsheetImporter(DataSet data) {
            _data = data;
            // Set the defaults for the validation/confirmation checkers
            RowConfirmPercentage = 0;
            ValidateConfirmPercentage = 0;
        }

        public Dictionary<int, string> TableInfo() {
            var retval = new Dictionary<int, string>();

            for (int i = 0; i < _data.Tables.Count; i++) {
                retval.Add(i, _data.Tables[i].ToString());
            }

            return retval;
        }

        public IEnumerable<IPowersheetPropertyMap> GetMappings(params string[] propertyNames) {
            var retval = new List<IPowersheetPropertyMap>();
            PropertyInfo[] properties = typeof(T).GetProperties();

            foreach (var property in properties) {
                var attribute = Attribute.GetCustomAttribute(property, typeof(SpreadsheetColumn)) as SpreadsheetColumn;
                if (((propertyNames == null || propertyNames.Count() == 0) || propertyNames.Contains(property.Name)) && attribute != null) {
                    foreach (string mapValue in attribute.SpreadsheetMapValues) {
                        var map = new PowersheetPropertyMap() {
                            PropertyName = property.Name,
                            IsValueRequired = attribute.RequiredProperty,
                            SpreadsheetMapValue = mapValue
                        };
                        retval.Add(map);
                    }
                }
            }

            return retval;
        }

        public IEnumerable<T> GetAll(int tableId) {
            return Retrieve(tableId, null, null, null, null);
        }

        public IEnumerable<T> GetAll(int tableId, int headingsRowIndex) {
            return Retrieve(tableId, headingsRowIndex, null, null, null);
        }

        public IEnumerable<T> Fetch(int tableId, int? headingsRowIndex, int? start, int? limit, params IPowersheetPropertyMap[] selectedColumns) {
            return Retrieve(tableId, headingsRowIndex, start, limit, selectedColumns);
        }

        public IEnumerable<IPowersheetDump> Extract(int tableId, int headingsRowIndex, params string[] ignoreColumns) {
            // There may be a better way to execute this method, but get it working!
            if (tableId >= _data.Tables.Count) {
                throw new Exception("Table index is out with the range of the data set");
            }

            if (headingsRowIndex >= _data.Tables[tableId].Rows.Count || ((headingsRowIndex + 1) >= _data.Tables[tableId].Rows.Count)) {
                throw new Exception("Headings Row specified is out with the range of the data set");
            }

            DataTable currentTable = _data.Tables[tableId];
            var headingsRow = new Dictionary<int, string>();

            // Get headings row data
            object[] headingsObjRow = currentTable.Rows[headingsRowIndex].ItemArray;
            for (int i = 0; i < headingsObjRow.Count(); i++) {
                object v = headingsObjRow[i];
                try {
                    string vString = v.ToString();
                    if (!ignoreColumns.Contains(vString)) {
                        headingsRow.Add(i, vString);
                    }
                }
                catch {
                    continue;
                }
            }

            var retval = new List<IPowersheetDump>();
            for (int i = (headingsRowIndex + 1); i < currentTable.Rows.Count; i++) {
                T newObj = (T)Activator.CreateInstance(typeof(T));

                var cols = new Dictionary<string, string>();
                object[] rowData = currentTable.Rows[i].ItemArray;

                foreach (var pair in headingsRow) {
                    // Find out if the key is within rowdata, and ALSO if the property exists, if not; then put it in cols
                    if (pair.Key < rowData.Length) {
                        if (typeof(T).HasProperty(pair.Value)) {

                        }
                        else {

                        }
                    }


                    if (pair.Key < rowData.Length) {
                        object rowObject = rowData[pair.Key];
                        if (rowObject != null) {
                            try {
                                string rowString = rowObject.ToString();
                            }
                            catch {
                                continue;
                            }
                        }
                    }
                }
            }

            return retval;
        }

        public object[,] ToGrid(int tableId) {
            if (tableId >= _data.Tables.Count) {
                throw new Exception("Table index is out with the range of the data set");
            }

            int width = _data.Tables[tableId].MaxWidth();
            int height = _data.Tables[tableId].Rows.Count;
            // [y,x]
            var retval = new object[height, width];

            for (int o = 0; o < _data.Tables[tableId].Rows.Count; o++) {
                object[] row = _data.Tables[tableId].Rows[o].ItemArray;
                for (int i = 0; i < row.Count(); i++) {
                    retval[o, i] = row[i];
                }
            }

            return retval;
        }

        private bool Validate(T newObj, IEnumerable<IPowersheetPropertyMap> columns, int propertiesMapped) {
            double percentageMapped = (double)(propertiesMapped * 100 / columns.Count());
            if (Math.Floor(percentageMapped) < ValidateConfirmPercentage) {
                return false;
            }

            foreach (var column in columns) {
                PropertyInfo p = typeof(T).GetProperty(column.PropertyName);
                if (p != null) {
                    object value;
                    if (p.PropertyType == typeof(string)) {
                        value = p.GetValue(newObj, null);
                        if (value != null) {
                            string vString = p.GetValue(newObj, null).ToString().Trim();
                            value = (String.IsNullOrWhiteSpace(vString)) ? null : vString;
                        }
                    }
                    else {
                        value = p.GetValue(newObj, null);
                    }

                    if (column.IsValueRequired && value == null) {
                        return false;
                    }
                }
            }

            return true;
        }

        private int GetHeadingsRowIndex(DataTable table, IEnumerable<IPowersheetPropertyMap> columns, int confirmPercentage = 70) {
            for (int i = 0; i < table.Rows.Count; i++) {
                int counter = 0;
                DataRow row = table.Rows[i];

                foreach (var column in columns) {
                    int x = row.ColumnIndexOf(column.SpreadsheetMapValue);
                    if (x >= 0) {
                        counter++;
                    }
                }

                double percentage = (double)(counter * 100) / columns.Count();
                if (Math.Floor(percentage) >= confirmPercentage) {
                    return i;
                }
            }

            return -1;
        }
 
        private IEnumerable<T> Retrieve(int tableId, int? headingsRowIndex, int? start, int? limit, params IPowersheetPropertyMap[] selectedColumns) {
            // Validate our table index
            if (tableId >= _data.Tables.Count) {
                throw new Exception("Table index is out with the range of the data set");
            }

            DataTable currentTable = _data.Tables[tableId];
            // Validate/Get our Columns/Property mappings
            List<IPowersheetPropertyMap> columns;
            if (selectedColumns == null || selectedColumns.Length == 0) {
                // Fetch the columns
                columns = GetMappings().ToList();
            }
            else {
                columns = selectedColumns.ToList();
            }

            // Find our Headings Row/Index
            headingsRowIndex = (headingsRowIndex != null) ? (int)headingsRowIndex : GetHeadingsRowIndex(currentTable, columns);
            if (headingsRowIndex < 0 || headingsRowIndex >= currentTable.Rows.Count) {
                throw new Exception("Headings row index is out with the range of the data set");
            }

            // Get the Start and End of extraction
            int iStart = (start != null) ? ((int)headingsRowIndex + (int)start) : ((int)headingsRowIndex + 1);
            int iLimit = (limit != null && ((int)headingsRowIndex + (int)limit < currentTable.Rows.Count))
                ? ((int)headingsRowIndex + (int)limit)
                : currentTable.Rows.Count;

            // Begin extraction
            var retval = new List<T>();
            int[] columnIndexes = columns.ColumnIndexes();
            DataRow headingsRow = currentTable.Rows[(int)headingsRowIndex];

            for (var i = iStart; i < iLimit; i++) {
                DataRow currentRow = currentTable.Rows[i];

                int nonNullColumns = currentRow.NonNullColumnCount(columnIndexes);
                if (nonNullColumns <= 0) {
                    continue;
                }

                double nonNullPercentage = (double)(nonNullColumns * 100) / columns.Count();
                if (Math.Floor(nonNullPercentage) <= RowConfirmPercentage) {
                    continue;
                }

                int propertiesMapped = 0;
                T newObj = (T)Activator.CreateInstance(typeof(T));

                foreach (var column in columns) {
                    try {
                        PropertyInfo p = typeof(T).GetProperty(column.PropertyName);
                        if (p != null) {
                            if (column.SpreadsheetMapValue.IsFormula()) {
                                IPowersheetFormula formula = PowersheetFormulaFactory.Get(column.SpreadsheetMapValue);
                                if (formula != null) {
                                    object value = formula.Execute(ref _data, ref p, tableId, (int)headingsRowIndex, i);
                                    if (value != null) {
                                        p.SetValue(newObj, value, null);
                                        propertiesMapped++;
                                    }
                                }
                            }
                            else {
                                int index = headingsRow.ColumnIndexOf(column.SpreadsheetMapValue);
                                if (index >= 0) {
                                    object value = currentRow[index].CastToPropertyType(p);
                                    if (value != null) {
                                        p.SetValue(newObj, value, null);
                                        propertiesMapped++;
                                    }
                                }
                            }
                        }
                    }
                    catch {
                        continue;
                    }
                }

                if (Validate(newObj, columns, propertiesMapped)) {
                    retval.Add(newObj);
                }
            }
            return retval;
        }
    }
}
