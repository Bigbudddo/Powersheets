using Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {
    
    public sealed class PowersheetImporter<T> where T : class {

        private bool _canRun = false;
        private DataSet _data;
        private int _confirmPercentage;
        private string _fileName;
        private string _fileExtension;

        public bool CanRun {
            get {
                return _canRun;
            }
        }

        public DataSet DataSet {
            get {
                return _data;
            }
        }

        public string FileName {
            get {
                return _fileName;
            }
        }

        public string FileExtension {
            get {
                return _fileExtension;
            }
        }

        public PowersheetImporter(string fileName) {
            _fileName = fileName;
            _fileExtension = fileName.ExtractFileExtension();
            _data = this.FetchDataSet(fileName);

            if (_data != null) {
                _canRun = true;
            }
        }

        public PowersheetImporter(ref DataSet data) {
            _fileName = "Not Supplied";
            _fileExtension = "Not Supplied";
            _data = data;

            if (_data != null) {
                _canRun = true;
            }
        }

        public PowersheetImporter(ref Stream dataStream, string fileName) {
            _fileName = fileName;
            _fileExtension = fileName.ExtractFileExtension();
            _data = FetchDataSet(ref dataStream, fileName);

            if (_data != null) {
                _canRun = true;
            }
        }

        // TODO:
        public DataRow GetHeadingsRow(int tableId, int percentage, params string[] columnHeadings) {
            throw new NotImplementedException();
        }

        // TODO:
        public int GetHeadingsRowIndex(int tableId, int percentage, params string[] columnHeadings) {
            throw new NotImplementedException();
        }

        public int GetTableId(string tableName) {
            if (!_canRun) {
                return -1;
            }

            for (int i = 0; i < _data.Tables.Count; i++) {
                string tableValue = _data.Tables[i].ToString();
                if (tableValue == tableName) {
                    return i;
                }
            }

            return -1;
        }

        public string GetTableName(int tableId) {
            if (!_canRun || tableId >= _data.Tables.Count) {
                return null;
            }
            return _data.Tables[tableId].ToString();
        }

        public IEnumerable<KeyValuePair<int, string>> GetTableInformation() {
            if (!_canRun) {
                return null;
            }

            var retval = new Dictionary<int, string>();

            for (int i = 0; i < _data.Tables.Count; i++) {
                retval.Add(i, _data.Tables[i].ToString());
            }

            return retval;
        }

        public IEnumerable<IPowersheetPropertyMap> GetPropertyMappings(params string[] propertyNames) {
            var retval = new List<PowersheetPropertyMap>();
            PropertyInfo[] properties = typeof(T).GetProperties();

            foreach (var property in properties) {
                var attribute = Attribute.GetCustomAttribute(property, typeof(SpreadsheetColumn)) as SpreadsheetColumn;
                if (((propertyNames == null || propertyNames.Count() == 0) || propertyNames.Contains(property.Name)) && attribute != null) {
                    foreach (string mapValue in attribute.SpreadsheetMapValues) {
                        var map = new PowersheetPropertyMap() {
                            PropertyName = property.Name,
                            IsValueRequired = attribute.RequiredProperty,
                            SpreadsheetMapValue = mapValue,
                        };
                        retval.Add(map);
                    }
                }
            }

            return retval;
        }

        public IEnumerable<T> GetAll(int tableId, int headingsRowIndex, params IPowersheetPropertyMap[] columns) {
            return Fetch(tableId, headingsRowIndex, null, null, columns);   
        }

        public IEnumerable<T> Fetch(int tableId, int? headingsRowIndex, int? start, int? limit, params IPowersheetPropertyMap[] columns) {
            if (!_canRun || tableId >= _data.Tables.Count) {
                return null;
            }
            
            headingsRowIndex = (headingsRowIndex != null) ? (int)headingsRowIndex : GetHeadingsRowIndex(tableId, 60, columns.ColumnHeadings(true));
            if (headingsRowIndex < 0 || headingsRowIndex >= _data.Tables[tableId].Rows.Count) {
                return null;
            }

            int iStart = (start != null) ? ((int)headingsRowIndex + (int)start) : ((int)headingsRowIndex + 1);
            int iLimit = (limit != null && ((int)headingsRowIndex + (int)limit < _data.Tables[tableId].Rows.Count))
                ? ((int)headingsRowIndex + (int)limit) 
                : _data.Tables[tableId].Rows.Count;

            var retval = new List<T>();
            int[] columnIndexes = columns.ColumnIndexes();
            DataTable currentTable = _data.Tables[tableId];
            DataRow headingsRow = _data.Tables[tableId].Rows[(int)headingsRowIndex];

            for (var i = iStart; i < iLimit; i++) {
                DataRow currentRow = currentTable.Rows[i];

                int nonNullColumns = currentRow.NonNullColumnCount(columnIndexes);
                if (nonNullColumns <= 0) {
                    continue;
                }

                double nonNullPercentage = (double)(nonNullColumns * 100) / columns.Count();
                if (Math.Floor(nonNullPercentage) <= _confirmPercentage) {
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

                if (this.Validate(newObj, columns, propertiesMapped)) {
                    retval.Add(newObj);
                }
            }

            return retval;
        }

        public object[,] FetchGrid(int tableId) {
            if (!_canRun || tableId >= _data.Tables.Count) {
                return null;
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
            if (Math.Floor(percentageMapped) < _confirmPercentage) {
                return false;
            }

            foreach (var column in columns) {
                PropertyInfo p = typeof(T).GetProperty(column.PropertyName);
                if (p != null) {
                    object value;
                    if (p.PropertyType == typeof(string)) {
                        string vString = p.GetValue(newObj, null).ToString().Trim();
                        value = (String.IsNullOrWhiteSpace(vString)) ? null : vString;
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

        private int CountMapableProperties(params string[] propertyNames) {
            int retval = 0;
            PropertyInfo[] properties = typeof(T).GetProperties();

            foreach (var property in properties) {
            }

            return retval;
        }

        private DataSet FetchDataSet(string fileName) {
            if (string.IsNullOrWhiteSpace(fileName)) {
                return null;
            }
            
            string fileExtension = fileName.ExtractFileExtension();
            try {
                using (FileStream stream = File.Open(fileName, FileMode.Open)) {
                    switch (fileExtension.ToUpper()) {
                        case "XLS":
                            return ExcelReaderFactory.CreateBinaryReader(stream).AsDataSet();
                        case "XLSX":
                            return ExcelReaderFactory.CreateOpenXmlReader(stream).AsDataSet();
                        default:
                            return null;
                    }
                }
            }
            catch {
                return null;
            }
        }

        private DataSet FetchDataSet(ref Stream stream, string fileName) {
            if (stream == null || !stream.CanRead) {
                return null;
            }

            string fileExtension = fileName.ExtractFileExtension();
            try {
                switch (fileExtension.ToUpper()) {
                    case "XLS":
                        return ExcelReaderFactory.CreateBinaryReader(stream).AsDataSet();
                    case "XLSX":
                        return ExcelReaderFactory.CreateOpenXmlReader(stream).AsDataSet();
                    default:
                        return null;
                }
            }
            catch {
                return null;
            }
        }
    }
}
