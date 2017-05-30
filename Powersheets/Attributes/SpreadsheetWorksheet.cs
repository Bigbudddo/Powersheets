using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {

    [AttributeUsage(AttributeTargets.Property)]
    public sealed class SpreadsheetWorksheet : Attribute {

        public readonly string WorksheetTitle;

        public SpreadsheetWorksheet(string title) {
            WorksheetTitle = title;
        }
    }
}
