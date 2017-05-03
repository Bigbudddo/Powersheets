using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {
    
    public static class PowersheetExportFactory {

        // TODO: add code back in once the XLS exporting has been completed/fixed

        //public enum ExportType {
        //    CSV, XLS
        //}

        //public static IPowersheetExporter Get(ExportType type) {
        //    switch (type) {
        //        case ExportType.CSV:
        //            return new CSVExporter();
        //        case ExportType.XLS:
        //            return new XLSExporter();
        //        default:
        //            return null;
        //    }
        //}

        public static IPowersheetExporter Get() {
            return new CSVExporter();
        }
    }
}
