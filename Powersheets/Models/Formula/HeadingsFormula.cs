using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {
    
    internal sealed class HeadingsFormula : PowersheetFormula, IPowersheetFormula {

        public HeadingsFormula(string formula)
            : base(formula) {
            Formula = formula;
        }

        public object Execute(ref DataSet dataSet, ref PropertyInfo property, int currTableId, int currTableHeadingsRow, int currRowIndex) {
            string subType = Formula.ExtractFormulaSubType();

            switch (subType) {
                case "EQUALS":
                    return ExecuteEquals(ref dataSet, ref property, currTableId, currTableHeadingsRow, currRowIndex);
                case "LIKE":
                    return ExecuteLike(ref dataSet, ref property, currTableId, currTableHeadingsRow, currRowIndex);
                default:
                    return null;
            }
        }

        private object ExecuteEquals(ref DataSet dataSet, ref PropertyInfo property, int currTableId, int currTableHeadingsRow, int currRowIndex) {
            string value = Formula.ExtractFormulaValue().ToLower().Trim();
            string cast = Formula.ExtractFormulaCastType();

            int counter = 0;
            int toFindBeforeReturn;
            if (int.TryParse(cast, out toFindBeforeReturn)) {
                for (int i = 0; i < dataSet.Tables[currTableId].Rows[currTableHeadingsRow].ItemArray.Count(); i++) {
                    object obj = dataSet.Tables[currTableId].Rows[currTableHeadingsRow].ItemArray[i];
                    try {
                        string objString = obj.ToString().ToLower().Trim();
                        if (value == objString) {
                            counter++;

                            if (counter >= toFindBeforeReturn) {
                                // We have the column index, get the value!
                                return dataSet.Tables[currTableId].Rows[currRowIndex].ItemArray[i].CastToPropertyType(property);
                            }
                        }
                    }
                    catch {
                        continue;
                    }
                }
            }
            return null;
        }

        private object ExecuteLike(ref DataSet dataSet, ref PropertyInfo property, int currTableId, int currTableHeadingsRow, int currRowIndex) {
            string value = Formula.ExtractFormulaValue().ToLower().Trim();
            string cast = Formula.ExtractFormulaCastType();

            int maxComputeLength;
            if (int.TryParse(cast, out maxComputeLength)) {
                for (int i = 0; i < dataSet.Tables[currTableId].Rows[currTableHeadingsRow].ItemArray.Count(); i++) {
                    object obj = dataSet.Tables[currTableId].Rows[currTableHeadingsRow].ItemArray[i];
                    try {
                        string objString = obj.ToString().ToLower().Trim();
                        int compute = value.LevenshteinCompute(objString);

                        if (compute <= maxComputeLength) {
                            // Get the value
                            return dataSet.Tables[currTableId].Rows[currRowIndex].ItemArray[i].CastToPropertyType(property);
                        }
                    }
                    catch {
                        continue;
                    }
                }
            }
            return null;
        }
    }
}
