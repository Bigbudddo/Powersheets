using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {
    
    internal sealed class SelectFormula : PowersheetFormula, IPowersheetFormula {

        public SelectFormula(string formula) : base(formula) {
            Formula = formula;
        }

        public object Execute(ref DataSet dataSet, ref PropertyInfo property, int currTableId, int currTableHeadingsRow, int currRowIndex) {
            string subType = Formula.ExtractFormulaSubType();

            switch (subType) {
                case "CELL":
                    return ExecuteCell(ref dataSet, ref property, currTableId, currTableHeadingsRow, currRowIndex);
                case "LOOK":
                    return ExecuteLook(ref dataSet, ref property, currTableId, currTableHeadingsRow, currRowIndex);
                case "WORKBOOK":
                    return ExecuteWorkbook(ref dataSet, ref property, currTableId, currTableHeadingsRow, currRowIndex);
                case "WORKLOOK":
                    return ExecuteWorklook(ref dataSet, ref property, currTableId, currTableHeadingsRow, currRowIndex);
                default:
                    return null;
            }
        }

        private object ExecuteCell(ref DataSet dataSet, ref PropertyInfo property, int currTableId, int currTableHeadingsRow, int currRowIndex) {
            string value = Formula.ExtractFormulaValue();
            string[] values = value.Split(',');

            int x;
            int y;
            if (int.TryParse(values[0], out x) && int.TryParse(values[1], out y)) {
                if (y >= 0 && y < dataSet.Tables[currTableId].Rows.Count) {
                    object[] rowData = dataSet.Tables[currTableId].Rows[y].ItemArray;

                    if (x >= 0 && x < rowData.Length) {
                        if (rowData[x] != null && rowData[x].GetType() != typeof(DBNull)) {
                            return rowData[x].CastToPropertyType(property);
                        }
                    }
                }
            }
            return null;
        }

        private object ExecuteLook(ref DataSet dataSet, ref PropertyInfo property, int currTableId, int currTableHeadingsRow, int currRowIndex) {
            string value = Formula.ExtractFormulaValue();
            string[] values = value.Split(',');

            if (values.Length == 3) {
                
                int x;
                int y;
                string lookup = values[0].ToLower().Trim();
                if (int.TryParse(values[1], out x) && int.TryParse(values[2], out y)) {
                    for (int o = 0; o < dataSet.Tables[currTableId].Rows.Count; o++) {
                        object[] rowData = dataSet.Tables[currTableId].Rows[o].ItemArray;
                        
                        for (int i = 0; i < rowData.Length; i++) {
                            if (rowData[i] != null && rowData[i].GetType() != typeof(DBNull)) {
                                string rowString = rowData[i].ToString().ToLower().Trim();
                                if (lookup == rowString) {
                                    int offX = (i + x);
                                    int offY = (o + y);

                                    if (offY >= 0 && offY < dataSet.Tables[currTableId].Rows.Count) {
                                        object[] offRowData = dataSet.Tables[currTableId].Rows[offY].ItemArray;
                                        
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
            return null;
        }
    
        private object ExecuteWorkbook(ref DataSet dataSet, ref PropertyInfo property, int currTableId, int currTableHeadingsRow, int currRowIndex) {
            string value = Formula.ExtractFormulaValue();
            string[] values = value.Split(',');

            if (values.Length == 3) {

                int x;
                int y;
                int workbookIndex = dataSet.TableIndexOf(values[0]);
                if (workbookIndex >= 0 && int.TryParse(values[1], out x) && int.TryParse(values[2], out y)) {
                    if (y >= 0 && y < dataSet.Tables[workbookIndex].Rows.Count) {
                        object[] rowData = dataSet.Tables[workbookIndex].Rows[y].ItemArray;
                        
                        if (x >= 0 && x < rowData.Length) {
                            if (rowData[x] != null && rowData[x].GetType() != typeof(DBNull)) {
                                return rowData[x].CastToPropertyType(property);
                            }
                        }
                    }
                }
            }
            return null;
        }

        private object ExecuteWorklook(ref DataSet dataSet, ref PropertyInfo property, int currTableId, int currTableHeadingsRow, int currRowIndex) {
            string value = Formula.ExtractFormulaValue();
            string[] values = value.Split(',');

            if (values.Length == 4) {

                int x;
                int y;
                int workbookIndex = dataSet.TableIndexOf(values[0]);
                string lookup = values[1].ToLower().Trim();
                if (workbookIndex >= 0 && int.TryParse(values[2], out x) && int.TryParse(values[3], out y)) {
                    for (int o = 0; o < dataSet.Tables[workbookIndex].Rows.Count; o++) {
                        object[] rowData = dataSet.Tables[workbookIndex].Rows[o].ItemArray;

                        for (int i = 0; i < rowData.Length; i++) {
                            if (rowData[i] != null && rowData[i].GetType() != typeof(DBNull)) {
                                string rowString = rowData[i].ToString().ToLower().Trim();
                                if (lookup == rowString) {
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
            return null;
        }
    }
}
