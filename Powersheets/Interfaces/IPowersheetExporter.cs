using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {
    
    public interface IPowersheetExporter {

        StringBuilder Dump(IEnumerable<IPowersheetDump> dataSet, bool writeHeadings, bool writeAutoIncrement);

        StringBuilder Dump(IEnumerable<IPowersheetDump> dataSet, IEnumerable<string> propertyColumns, bool writeHeadings, bool writeAutoIncrement);

        StringBuilder Export(IEnumerable<object> dataSet, bool writeHeadings, bool writeAutoIncrement);

        StringBuilder Export(IEnumerable<object> dataSet, IEnumerable<string> columns, bool writeHeadings, bool writeAutoIncrement);
    }
}
