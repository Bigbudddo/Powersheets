using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {

    public sealed class XLSExporter : Exporter, IPowersheetExporter {

        private string XLSHeading {
            get {
                var sb = new StringBuilder();

                sb.Append("<?xml version=\"1.0\"?>\n");
                sb.Append("<?mso-application progid=\"Excel.Sheet\"?>\n");
                sb.Append("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" ");
                sb.Append("xmlns:o=\"urn:schemas-microsoft-com:office:office\" ");
                sb.Append("xmlns:x=\"urn:schemas-microsoft-com:office:excel\" ");
                sb.Append("xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" ");
                sb.Append("xmlns:html=\"http://www.w3.org/TR/REC-html40\">\n");
                sb.Append("<DocumentProperties xmlns=\"urn:schemas-microsoft-com:office:office\">");
                sb.Append("<Author>Stuart Harrison</Author>");
                sb.Append("</DocumentProperties>");
                sb.Append("<ExcelWorkbook xmlns=\"urn:schemas-microsoft-com:office:excel\">\n");
                sb.Append("<ProtectStructure>False</ProtectStructure>\n");
                sb.Append("<ProtectWindows>False</ProtectWindows>\n");
                sb.Append("</ExcelWorkbook>\n");

                return sb.ToString();
            }
        }

        private string XLSStyles {
            get {
                var sb = new StringBuilder();

                sb.Append("<Styles>\n");
                sb.Append("<Style ss:ID=\"Default\" ss:Name=\"Normal\">\n");
                sb.Append("<Alignment ss:Vertical=\"Bottom\"/>\n");
                sb.Append("<Borders/>\n");
                sb.Append("<Font/>\n");
                sb.Append("<Interior/>\n");
                sb.Append("<NumberFormat/>\n");
                sb.Append("<Protection/>\n");
                sb.Append("</Style>\n");
                sb.Append("<Style ss:ID=\"s27\" ss:Name=\"Hyperlink\">\n");
                sb.Append("<Font ss:Color=\"#0000FF\" ss:Underline=\"Single\"/>\n");
                sb.Append("</Style>\n");
                sb.Append("<Style ss:ID=\"s24\">\n");
                sb.Append("<Font x:Family=\"Swiss\" ss:Bold=\"1\"/>\n");
                sb.Append("</Style>\n");
                sb.Append("<Style ss:ID=\"s25\">\n");
                sb.Append("<Font x:Family=\"Swiss\" ss:Italic=\"1\"/>\n");
                sb.Append("</Style>\n");
                sb.Append("<Style ss:ID=\"s26\">\n");
                sb.Append("<Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Bottom\"/>\n");
                sb.Append("</Style>\n");
                sb.Append("</Styles>\n");

                return sb.ToString();
            }
        }

        private string XLSOptions {
            get {
                // This is Required Only Once ,	But this has to go after the First Worksheet's First Table	
                var sb = new StringBuilder();
                sb.Append("\n<WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\">\n<Selected/>\n </WorksheetOptions>\n");
                return sb.ToString();
            }
        }

        public StringBuilder Export(IEnumerable<object> dataSet, bool writeHeadings, bool writeAutoIncrement) {
            if (dataSet == null || dataSet.Count() <= 0) {
                throw new Exception("Dataset cannot be null/or empty");
            }

            var type = dataSet.ToList().GetType().GetTypeInfo().GenericTypeArguments[0];
            IEnumerable<string> columns = FetchObjectProperties(type, null);

            return Export(dataSet, columns, writeHeadings, writeAutoIncrement);
        }

        public StringBuilder Export(IEnumerable<object> dataSet, IEnumerable<string> columns, bool writeHeadings, bool writeAutoIncrement) {
            var sb = new StringBuilder();

            if (dataSet == null || dataSet.Count() <= 0) {
                throw new Exception("Dataset cannot be null/or empty");
            }

            string title = "Sheet 1";

            var type = dataSet.ToList().GetType().GetTypeInfo().GenericTypeArguments[0];
            var attribute = Attribute.GetCustomAttribute(type, typeof(SpreadsheetWorksheet)) as SpreadsheetWorksheet;
            if (attribute != null) {
                title = attribute.WorksheetTitle;
            }

            sb.Append(this.XLSHeading);
            sb.Append(this.XLSStyles);
            sb.Append(this.GenerateWorksheet(dataSet, columns, title, writeHeadings, writeAutoIncrement).ToString());
            sb.Append(this.XLSOptions);
            sb.Append("</Workbook>\n");

            return sb.ConvertHtmlToXLS();
        }

        public StringBuilder ExportMultiple(IEnumerable<string> workbookNames, bool writeHeadings, bool writeAutoIncrement, params IEnumerable<object>[] workbooks) {
            var sb = new StringBuilder();

            if (workbooks == null || workbooks.Count() <= 0) {
                throw new Exception("Dataset(s) cannot be null/or empty");
            }

            if (workbookNames.Count() != workbooks.Count()) {
                throw new Exception("Workbook names does not have the same length as the data!");
            }

            sb.Append(this.XLSHeading);
            sb.Append(this.XLSStyles);
            sb.Append(this.GenerateWorksheet(workbooks.First(), null, workbookNames.First(), writeHeadings, writeAutoIncrement).ToString());
            sb.Append(this.XLSOptions);

            for (int i = 1; i < workbooks.Count(); i++) {
                sb.Append(this.GenerateWorksheet(workbooks[i], null, workbookNames.ToList()[i], writeHeadings, writeAutoIncrement).ToString());
            }

            sb.Append("</Workbook>\n");

            return sb.ConvertHtmlToXLS();
        }

        public StringBuilder PushDump(IEnumerable<IPowersheetDump> dataSet, bool writeHeadings, bool writeAutoIncrement) {
            throw new NotImplementedException();
        }

        public StringBuilder PushDump(IEnumerable<IPowersheetDump> dataSet, IEnumerable<string> propertyColumns, bool writeHeadings, bool writeAutoIncrement) {
            throw new NotImplementedException();
        }

        private StringBuilder GenerateWorksheet(IEnumerable<object> data, IEnumerable<string> propertyColumns, string title, bool writeHeadings, bool writeAutoIncrement) {
            var sb = new StringBuilder();
            var dataList = data.ToList();

            var type = dataList.GetType().GetTypeInfo().GenericTypeArguments[0];
            List<string> columnCollection = new List<string>();

            if (writeHeadings) {
                var attribute = Attribute.GetCustomAttribute(type, typeof(SpreadsheetAutoIncrement)) as SpreadsheetAutoIncrement;
                if (attribute != null) {
                    columnCollection.Add(attribute.PropertyName);
                }
            }

            if (propertyColumns == null || propertyColumns.Count() <= 0) {
                columnCollection.AddRange(FetchObjectProperties(type, null));
            }
            else {
                columnCollection.AddRange(propertyColumns);
            }
            

            sb.Append("<Worksheet ss:Name=\"" + title + "\">");
            sb.Append("<Table>");

            if (writeHeadings) {
                sb.Append("<tr>");
                foreach (var i in columnCollection) {
                    sb.Append("<th>");
                    sb.Append(i);
                    sb.Append("</th>");
                }
                sb.Append("</tr>");
            }

            for (int r = 0; r < data.Count(); r++) {
                // foreach row in the data
                var item = dataList[r];
                sb.Append("<tr>");

                foreach (var c in columnCollection) {
                    sb.Append("<td>");
                    try {
                        var propValue = item.GetType().GetProperty(c).GetValue(item, null);
                        string value = propValue.ToStringValue();

                        sb.Append(value);
                    }
                    catch {

                    }
                    sb.Append("</td>");
                }

                sb.Append("</tr>");
            }

            sb.Append("</Table>");
            sb.Append("</Worksheet>");
            return sb;
        }
    }
}
