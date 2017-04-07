using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {
    
    public interface IPowersheetFormula {

        object Execute(ref DataSet dataSet, ref PropertyInfo property, int currTableId, int currTableHeadingsRow, int currRowIndex);
    }
}
