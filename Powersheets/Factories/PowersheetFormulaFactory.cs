using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {
    
    internal static class PowersheetFormulaFactory {

        public static IPowersheetFormula Get(string formula) {
            if (String.IsNullOrWhiteSpace(formula) || !formula.IsFormula()) {
                return null;
            }

            string formulaType = formula.ExtractFormulaType();
            switch (formulaType) {
                case "CAST":
                    return new CastFormula(formula);
                case "SELECT":
                    return new SelectFormula(formula);
                case "MATH":
                    return new MathFormula(formula);
                case "VALUE":
                    return new ValueFormula(formula);
                case "HEADING":
                    return new HeadingsFormula(formula);
                case "RESULT":
                    return new ResultFormula(formula);
                default:
                    return null;
            }
        }
    }
}
