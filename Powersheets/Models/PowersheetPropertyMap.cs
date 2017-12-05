using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {
    
    internal sealed class PowersheetPropertyMap : IPowersheetPropertyMap {

        public bool IsValueRequired { get; set; }
        public int ColumnIndex { get; set; }
        public string PropertyName { get; set; }
        public string SpreadsheetMapValue { get; set; }
    }
}
