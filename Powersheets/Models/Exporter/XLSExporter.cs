using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {
    
    internal sealed class XLSExporter : IPowersheetExporter {

        public StringBuilder Export(IEnumerable<object> dataSet, IEnumerable<string> columns, bool writeHeadings, bool writeAutoIncrement) {
            var rowData = new StringBuilder();
            var colData = new StringBuilder();

            if (dataSet == null || dataSet.Count() <= 0) {
                return new StringBuilder();
            }

            List<string> columnCollection = new List<string>();
            if (writeAutoIncrement) {
                var attribute = Attribute.GetCustomAttribute(dataSet.First().GetType(), typeof(SpreadsheetAutoIncrement)) as SpreadsheetAutoIncrement;
                if (attribute != null) {
                    columnCollection.Add(attribute.PropertyName);
                }
            }
            columnCollection.AddRange(columns);

            rowData.Append("<Row ss:StyleID=\"s62\">");
            foreach (var s in columnCollection) {
                if (dataSet.First().GetType().GetProperty(s).GetType() == typeof(DateTime)) {
                    colData.Append("<Column ss:AutoFitWidth=\"1\" ss:Width=\"100\"/>");
                }
                else {
                    colData.Append("<Column ss:AutoFitWidth=\"1\" />");
                }
                if (writeHeadings) {
                    rowData.Append(String.Format("<Cell><Data ss:Type=\"String\" x:AutoFilter=\"All\">{0}</Data></Cell>", s));
                }
            }
            rowData.Append("</Row>");

            foreach (var data in dataSet) {
                rowData.Append("<Row>");
                foreach (string s in columnCollection) {
                    var propValue = data.GetType().GetProperty(s).GetValue(data, null);
                    string value = propValue.ToStringValue();
                    rowData.Append(String.Format("<Cell><Data ss:Type=\"String\">{0}</Data></Cell>", value));
                }
                rowData.Append("</Row>");
            }

            var dateHer = String.Format("export-{0}", DateTime.Now.ToShortDateString());
            StringBuilder retval = new StringBuilder(String.Format(
                @"<Worksheet ss:Name=""{0}"">
                    <Table ss:ExpandedColumnCount=""{1}"" ss:ExpandedRowCount=""{2}"" x:FullColumns=""1"" x:FullRows=""1"" ss:DefaultRowHeight=""15"">
                        {3}
                        {4}
                    </Table>
                    <WorksheetOptions xmlns=""urn:schemas-microsoft-com:office:excel"">
                        <PageSetup>
                            <Header x:Margin=""0.3""/>
                            <Footer x:Margin=""0.3""/>
                            <PageMargins x:Bottom=""0.75"" x:Left=""0.7"" x:Right=""0.7"" x:Top=""0.75""/>
                        </PageSetup>
                        <Print>
                            <ValidPrinterInfo/>
                            <HorizontalResolution>300</HorizontalResolution>
                            <VerticalResolution>300</VerticalResolution>
                        </Print>
                        <Selected/>
                        <Panes>
                            <Pane>
                                <Number>3</Number>
                                <ActiveCol>2</ActiveCol>
                            </Pane>
                        </Panes>
                        <ProtectObjects>False</ProtectObjects>
                        <ProtectScenarios>False</ProtectScenarios>
                    </WorksheetOptions>
                </Worksheet>", dateHer, columnCollection.Count() + 1, dataSet.Count() + 1, colData, rowData
            ));
            return retval;
        }
    }
}
