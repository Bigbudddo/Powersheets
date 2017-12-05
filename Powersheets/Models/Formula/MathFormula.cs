using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {
    
    internal sealed class MathFormula : PowersheetFormula, IPowersheetFormula {

        public MathFormula(string formula) : base(formula) {
            Formula = formula;
        }

        public object Execute(ref DataSet dataSet, ref PropertyInfo property, int currTableId, int currTableHeadingsRow, int currRowIndex) {
            string subType = Formula.ExtractFormulaSubType();

            switch (subType) {
                case "ADD":
                    return ExecuteMathAdd(ref dataSet, ref property, currTableId, currTableHeadingsRow, currRowIndex);
                case "SUBTRACT":
                    return ExecuteMathSubtract(ref dataSet, ref property, currTableId, currTableHeadingsRow, currRowIndex);
                default:
                    return null;
            }
        }

        public object ExecuteMathAdd(ref DataSet dataSet, ref PropertyInfo property, int currTableId, int currTableHeadingsRow, int currRowIndex) {
            string value = Formula.ExtractFormulaValue();
            string[] values = value.Split(',');

            decimal total = 0;
            DataRow headingsRow = dataSet.Tables[currTableId].Rows[currTableHeadingsRow];
            
            foreach (var v in values) {
                int indexOf = headingsRow.ColumnIndexOf(v);
                object[] dataRow = dataSet.Tables[currTableId].Rows[currRowIndex].ItemArray;

                if (indexOf >= 0 && indexOf < dataRow.Length) {
                    decimal d;
                    object data = dataSet.Tables[currTableId].Rows[currRowIndex].ItemArray[indexOf];

                    if (data != null && data.GetType() != typeof(DBNull)) {
                        if (decimal.TryParse(data.ToString(), out d)) {
                            total += d;
                        }
                    }
                }
            }

            return total.CastToPropertyType(property);
        }

        public object ExecuteMathSubtract(ref DataSet dataSet, ref PropertyInfo property, int currTableId, int currTableHeadingsRow, int currRowIndex) {
            bool initalFlagValue = true;
            string value = Formula.ExtractFormulaValue();
            string[] values = value.Split(',');

            decimal total = 0;
            DataRow headingsRow = dataSet.Tables[currTableId].Rows[currTableHeadingsRow];

            foreach (var v in values) {
                int indexOf = headingsRow.ColumnIndexOf(v);
                object[] dataRow = dataSet.Tables[currTableId].Rows[currRowIndex].ItemArray;

                if (indexOf >= 0 && indexOf < dataRow.Length) {
                    decimal d;
                    object data = dataSet.Tables[currTableId].Rows[currRowIndex].ItemArray[indexOf];

                    if (data != null && data.GetType() != typeof(DBNull)) {
                        if (decimal.TryParse(data.ToString(), out d)) {

                            if (initalFlagValue) {
                                total = d;
                                initalFlagValue = false;
                            }
                            else {
                                total -= d;
                            }
                        }
                    }
                }
            }

            return total.CastToPropertyType(property);
        }
    }
}
