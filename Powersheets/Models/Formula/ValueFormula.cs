using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {
    
    internal sealed class ValueFormula : PowersheetFormula, IPowersheetFormula {

        public ValueFormula(string formula) : base(formula) {
            Formula = formula;
        }

        public object Execute(ref DataSet dataSet, ref PropertyInfo property, int currTableId, int currTableHeadingsRow, int currRowIndex) {
            string subType = Formula.ExtractFormulaSubType();

            switch (subType) {
                case "SET":
                    return ExecuteSet(ref dataSet, ref property, currTableId, currTableHeadingsRow, currRowIndex);
                default:
                    return null;
            }
        }

        private object ExecuteSet(ref DataSet dataSet, ref PropertyInfo property, int currTableId, int currTableHeadingsRow, int currRowIndex) {
            string value = Formula.ExtractFormulaValue();
            return value.CastToPropertyType(property);
        }
    }
}
