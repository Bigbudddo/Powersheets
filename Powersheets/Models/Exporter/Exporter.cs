using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {

    internal abstract class Exporter {

        protected IEnumerable<string> FetchObjectProperties(Type objType, IEnumerable<string> columns) {
            var retval = new List<string>();
            PropertyInfo[] properties = objType.GetProperties();

            foreach (var property in properties) {
                var attribute = Attribute.GetCustomAttribute(property, typeof(SpreadsheetColumn)) as SpreadsheetColumn;
                if (((columns == null || columns.Count() == 0) || columns.Contains(property.Name)) && attribute != null) {
                    retval.Add(property.Name);
                }
            }

            return retval;
        }
    }
}
