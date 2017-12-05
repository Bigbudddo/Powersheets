using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {
    
    public interface IPowersheetPropertyMap {

        bool IsValueRequired { get; set; }
        int ColumnIndex { get; set; }
        string PropertyName { get; set; }
        string SpreadsheetMapValue { get; set; }
    }
}
