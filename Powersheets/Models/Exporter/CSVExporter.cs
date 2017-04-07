using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {
    
    internal sealed class CSVExporter : IPowersheetExporter {

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
