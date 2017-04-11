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

        IEnumerable<T> GetAll(int tableId);

        IEnumerable<T> GetAll(int tableId, int headingsRowIndex);

        IEnumerable<T> Fetch(int tableId, int? headingsRowIndex, int? start, int? limit, params IPowersheetPropertyMap[] selectedColumns);

        object[,] ToGrid(int tableId);
    }
}
