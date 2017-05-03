using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {
    
    public static class PowersheetExportFactory {

        public enum ExportType {
            CSV, XLS
        }

        public static IPowersheetExporter Get(ExportType type) {
            switch (type) {
                case ExportType.CSV:
                    return new CSVExporter();
                case ExportType.XLS:
                    throw new Exception("XLS Disabled");
                    //return new XLSExporter();
                default:
                    return null;
            }
        }
    }
}
