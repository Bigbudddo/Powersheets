using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {
    
    internal sealed class CSVExporter : Exporter, IPowersheetExporter {

        public StringBuilder PushDump(IEnumerable<IPowersheetDump> dataSet, bool writeHeadings, bool writeAutoIncrement) {
            if (dataSet == null || dataSet.Count() <= 0) {
                return new StringBuilder();
            }

            object dataObj = dataSet.First();
            IEnumerable<string> columns = FetchObjectProperties(dataObj.GetType(), null);

            return PushDump(dataSet, columns, writeHeadings, writeAutoIncrement);
        }

        public StringBuilder PushDump(IEnumerable<IPowersheetDump> dataSet, IEnumerable<string> propertyColumns, bool writeHeadings, bool writeAutoIncrement) {
            var builder = new StringBuilder();
            var rowBuilder = new StringBuilder();

            if (dataSet == null || dataSet.Count() <= 0) {
                return builder;
            }

            var columns = new List<string>();
            // Fetch the Auto-Increment Column if we need it/can get it
            var type = dataSet.First().GetType();
            var clsAttribute = Attribute.GetCustomAttribute(type, typeof(SpreadsheetAutoIncrement)) as SpreadsheetAutoIncrement;
            if (clsAttribute != null && writeAutoIncrement) {
                columns.Add(clsAttribute.PropertyName);
            }

            // Fetch the list of extra columns
            int longestIndex = dataSet.LargestIndexOf(); // Fetch the index out of the data
            List<string> probableHeadings = dataSet.ToArray()[longestIndex].Columns.Keys.ToList();
            // Do we need to fetch any additional columns, such as actual properties?
            if (propertyColumns != null && propertyColumns.Count() > 0) {
                columns.AddRange(propertyColumns);
            }

            // Check if we are writing headings out
            if (writeHeadings) {
                for (int i = 0; i < columns.Count; i++) {
                    rowBuilder.Append(String.Format("\"{0}\"", columns[i]));

                    if (i < (columns.Count() - 1)) {
                        rowBuilder.Append(",");
                    }
                }

                if (probableHeadings.Count() > 0 && rowBuilder.Length > 0) {
                    rowBuilder.Append(",");
                }

                for (int i = 0; i < probableHeadings.Count; i++) {
                    rowBuilder.Append(String.Format("\"{0}\"", probableHeadings[i]));

                    if (i < (probableHeadings.Count() - 1)) {
                        rowBuilder.Append(",");
                    }
                }

                builder.AppendLine(rowBuilder.ToString());
            }

            foreach (var row in dataSet) {
                rowBuilder.Clear();
                for (int i = 0; i < columns.Count; i++) {
                    var propValue = row.GetType().GetProperty(columns[i]).GetValue(row, null);
                    string value = propValue.ToString();

                    rowBuilder.Append(String.Format("\"{0}\"", value));

                    if (i < (columns.Count() - 1)) {
                        rowBuilder.Append(",");
                    } 
                }

                Dictionary<string, string> c = row.Columns;
                if (c.Count() > 0 && rowBuilder.Length > 0) {
                    rowBuilder.Append(",");
                }
                if (c.Count() > 0) {
                    for (int i = 0; i < c.Count; i++) {
                        string value = c.ElementAt(i).Value;
                        rowBuilder.Append(String.Format("\"{0}\"", value));

                        if (i < (c.Count() - 1)) {
                            rowBuilder.Append(",");
                        }
                    }
                }

                builder.AppendLine(rowBuilder.ToString());
            }

            return builder;
        }

        public StringBuilder Export(IEnumerable<object> dataSet, bool writeHeadings, bool writeAutoIncrement) {
            if (dataSet == null || dataSet.Count() <= 0) {
                return new StringBuilder();
            }

            object dataObj = dataSet.First();
            IEnumerable<string> columns = FetchObjectProperties(dataObj.GetType(), null);

            return Export(dataSet, columns, writeHeadings, writeAutoIncrement);
        }

        public StringBuilder Export(IEnumerable<object> dataSet, IEnumerable<string> columns, bool writeHeadings, bool writeAutoIncrement) {
            var builder = new StringBuilder();
            var rowBuilder = new StringBuilder();

            if (dataSet == null || dataSet.Count() <= 0) {
                return builder;
            }

            List<string> columnCollection = new List<string>();
            if (writeAutoIncrement) {
                var attribute = Attribute.GetCustomAttribute(dataSet.First().GetType(), typeof(SpreadsheetAutoIncrement)) as SpreadsheetAutoIncrement;
                if (attribute != null) {
                    columnCollection.Add(attribute.PropertyName);
                }
            }
            columnCollection.AddRange(columns);

            if (writeHeadings) {
                for (int i = 0; i < columnCollection.Count; i++) {
                    rowBuilder.Append(String.Format("\"{0}\"", columnCollection[i]));

                    if (i < columnCollection.Count() - 1) {
                        rowBuilder.Append(",");
                    }
                }
                builder.AppendLine(rowBuilder.ToString());
            }

            foreach (var data in dataSet) {
                rowBuilder.Clear();
                for (int i = 0; i < columnCollection.Count; i++) {
                    var propValue = data.GetType().GetProperty(columnCollection[i]).GetValue(data, null);
                    string value = propValue.ToStringValue();

                    rowBuilder.Append(String.Format("\"{0}\"", value));

                    if (i < columnCollection.Count() - 1) {
                        rowBuilder.Append(",");
                    }
                }
                builder.AppendLine(rowBuilder.ToString());
            }
            return builder;
        }
    }
}
