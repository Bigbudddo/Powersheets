using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {
    
    public interface IPowersheetExporter {

        StringBuilder Dump(IEnumerable<IPowersheetExporterDump> dataSet, bool writeHeadings);

        StringBuilder Export(IEnumerable<object> dataSet, IEnumerable<string> columns, bool writeHeadings, bool writeAutoIncrement);

        StringBuilder Export(IEnumerable<object> dataSet, bool writeHeadings, bool writeAutoIncrement);
    }
}
