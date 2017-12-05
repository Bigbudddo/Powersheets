using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {
    
    internal sealed class CastFormula : PowersheetFormula, IPowersheetFormula {

        public CastFormula(string formula) : base(formula) {
            Formula = formula;
        }

        public object Execute(ref DataSet dataSet, ref PropertyInfo property, int currTableId, int currTableHeadingsRow, int currRowIndex) {
            string subType = Formula.ExtractFormulaSubType();
            string value = Formula.ExtractFormulaValue();
            int indexOf = dataSet.Tables[currTableId].Rows[currTableHeadingsRow].ColumnIndexOf(value);

            if (indexOf >= 0) {
                return dataSet.Tables[currTableId].Rows[currRowIndex].ItemArray[indexOf].CastToStringType(subType);
            }
            return null;
        }
    }
}
