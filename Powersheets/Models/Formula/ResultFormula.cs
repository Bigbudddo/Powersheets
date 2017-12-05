using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {
    
    internal sealed class ResultFormula : PowersheetFormula, IPowersheetFormula {

        public ResultFormula(string formula) : base(formula) {
            Formula = formula;
        }

        public object Execute(ref DataSet dataSet, ref PropertyInfo property, int currTableId, int currTableHeadingsRow, int currRowIndex) {
            string subType = Formula.ExtractFormulaSubType();

            switch (subType) {
                case "WORKLOOK":
                    return ExecuteWorklook(ref dataSet, ref property, currTableId, currTableHeadingsRow, currRowIndex);
                default:
                    return null;
            }
        }

        private object ExecuteWorklook(ref DataSet dataSet, ref PropertyInfo property, int currTableId, int currTableHeadingsRow, int currRowIndex) {
            string value = Formula.ExtractFormulaValue();
            string[] values = value.Split(',');

            DataRow headingsRow = dataSet.Tables[currTableId].Rows[currTableHeadingsRow];

            if (values.Length == 4) {

                int x;
                int y;
                int workbookIndex = dataSet.TableIndexOf(values[0]);
                int indexOfHeading = headingsRow.ColumnIndexOf(values[1]);
                if (workbookIndex >= 0 && indexOfHeading >= 0 && int.TryParse(values[2], out x) && int.TryParse(values[3], out y)) {

                    object lookfor = dataSet.Tables[currTableId].Rows[currRowIndex].ItemArray[indexOfHeading];
                    if (lookfor != null && lookfor.GetType() != typeof(DBNull)) {

                        string lookForString = lookfor.ToString().ToLower().Trim();
                        for (int o = 0; o < dataSet.Tables[workbookIndex].Rows.Count; o++) {
                            object[] rowData = dataSet.Tables[workbookIndex].Rows[o].ItemArray;

                            for (int i = 0; i < rowData.Length; i++) {
                                if (rowData[i] != null && rowData[i].GetType() != typeof(DBNull)) {
                                    string rowString = rowData[i].ToString().ToLower().Trim();
                                    if (lookForString == rowString) {
                                        int offX = (i + x);
                                        int offY = (o + y);

                                        if (offY >= 0 && offY < dataSet.Tables[workbookIndex].Rows.Count) {
                                            object[] offRowData = dataSet.Tables[workbookIndex].Rows[offY].ItemArray;

                                            if (offX >= 0 && offX < offRowData.Length) {
                                                return offRowData[offX].CastToPropertyType(property);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return null;
        }
    }
}
