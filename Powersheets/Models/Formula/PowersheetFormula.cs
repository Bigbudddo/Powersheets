using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {
    
    internal abstract class PowersheetFormula {

        public string Formula { get; protected set; }

        public PowersheetFormula(string formula) {
            Formula = formula;
        }
    }
}
