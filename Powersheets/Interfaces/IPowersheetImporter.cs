using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Powersheets {

    public interface IPowersheetImporter<T> {

        DataSet Data { get; set; }

        int RowConfirmPercentage { get; set; }

        int ValidateConfirmPercentage { get; set; }

        Dictionary<int, string> TableInfo();

        IEnumerable<IPowersheetPropertyMap> GetMappings(params string[] propertyNames);

        IEnumerable<IPowersheetPropertyMap> GetIgnoreMappings(params string[] ignoreNames);

        IEnumerable<T> GetAll(int tableId);

        IEnumerable<T> GetAll(int tableId, int headingsRowIndex);

        IEnumerable<T> Fetch(int tableId, int? headingsRowIndex, int? start, int? limit, params IPowersheetPropertyMap[] selectedColumns);

        IEnumerable<T> PullDump(int tableId, int headingsRowIndex);

        object[,] ToGrid(int tableId);
    }

    public interface IPowersheetImporter {

        DataSet Data { get; set; }

        object[,] ToGrid(int tableId);

        Dictionary<int, string> TableInfo();

        IEnumerable<object[,]> DumpSpreadsheet();

        IEnumerable<Tuple<string, object[,]>> DumpSpreadsheetData();
    }
}
