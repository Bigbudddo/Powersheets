using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {
    
    internal sealed class CSVExporter : Exporter, IPowersheetExporter {

        public StringBuilder Dump(IEnumerable<IPowersheetExporterDump> dataSet, bool writeHeadings) {
            var builder = new StringBuilder();
            var rowBuilder = new StringBuilder();

            if (dataSet == null || dataSet.Count() <= 0) {
                return builder;
            }

            // Fetch the index of the longest dump collection..
            int longestIndex = dataSet.LargestIndexOf();

            List<string> probableHeadings = dataSet.ToArray()[longestIndex].Columns.Keys.ToList();
            if (writeHeadings) {
                for (int i = 0; i < probableHeadings.Count; i++) {
                    rowBuilder.Append(String.Format("\"{0}\"", probableHeadings[i]));

                    if (i < probableHeadings.Count() - 1) {
                        rowBuilder.Append(",");
                    }
                }
                builder.AppendLine(rowBuilder.ToString());
            }

            foreach (var row in dataSet) {
                rowBuilder.Clear();
                Dictionary<string, string> columns = row.Columns;
                for (int i = 0; i < columns.Count; i++) {
                    string value = columns.ElementAt(i).Value;
                    rowBuilder.Append(String.Format("\"{0}\"", value));

                    if (i < columns.Count() - 1) {
                        rowBuilder.Append(",");
                    }
                }
                builder.AppendLine(rowBuilder.ToString());
            }

            return builder;
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

        public StringBuilder Export(IEnumerable<object> dataSet, bool writeHeadings, bool writeAutoIncrement) {
            if (dataSet == null || dataSet.Count() <= 0) {
                return new StringBuilder();
            }

            object dataObj = dataSet.First();
            IEnumerable<string> columns = FetchObjectProperties(dataObj.GetType(), null);

            return Export(dataSet, columns, writeHeadings, writeAutoIncrement);
        }
    }
}
