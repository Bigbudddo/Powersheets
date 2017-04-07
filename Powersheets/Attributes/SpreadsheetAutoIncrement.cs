using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {
    
    [AttributeUsage(AttributeTargets.Class)]
    public sealed class SpreadsheetAutoIncrement : Attribute {

        public readonly string PropertyName;

        public SpreadsheetAutoIncrement(string propertyName) {
            PropertyName = propertyName;
        }
    }
}
