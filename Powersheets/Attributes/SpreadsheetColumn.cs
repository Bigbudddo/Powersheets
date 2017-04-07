using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {
    
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class SpreadsheetColumn : Attribute {

        public readonly bool RequiredProperty;
        public readonly List<string> SpreadsheetMapValues;

        public SpreadsheetColumn() {
            RequiredProperty = false;
            SpreadsheetMapValues = new List<string>();
        }

        public SpreadsheetColumn(bool required, params string[] mapValues) {
            RequiredProperty = required;
            SpreadsheetMapValues = mapValues.ToList();
        }

        public SpreadsheetColumn(params string[] mapValues) {
            RequiredProperty = false;
            SpreadsheetMapValues = mapValues.ToList();
        }
    }
}
