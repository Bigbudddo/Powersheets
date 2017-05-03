using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Powersheets;

namespace Powersheets.Console {
    class Program {

        static bool _run = true;

        static void Main(string[] args) {
            do {
                System.Console.Clear();
                System.Console.Write("Input: ");
                string input = System.Console.ReadLine();
                System.Console.WriteLine("-----------------");

                switch (input.ToUpper()) {
                    case "-1":
                        TestDump();
                        break;
                    case "1":
                        RunImport<MovieBase>(0, 0, null);
                        break;
                    case "2":
                        RunImport<MovieBaseHeadings>(0, 0, null);
                        break;
                    case "3":
                        RunImport<MovieBaseCategory>(1, 2, null);
                        break;
                    case "4":
                        RunImport<MovieBaseSet>(0, 0, null);
                        break;
                    case "5":
                        RunImport<MovieBaseCast>(0, 0, null);
                        break;
                    case "6":
                        RunImport<MovieBaseMath>(4, 0, null);
                        break;
                    case "7":
                        RunImport<MovieBaseCategoryResult>(1, 2, null);
                        break;
                    case "8":
                        RunImport<MovieBase>(0);
                        break;
                    case "9":
                        RunGrid<MovieBase>(0);
                        break;
                    case "Q":
                    case "EXIT":
                    case "QUIT":
                        _run = false;
                        break;
                }

                System.Console.WriteLine();
                System.Console.WriteLine("-----------------");
                System.Console.ReadKey();
            } while (_run);
        }

        static void RunGrid<T>(int tableId) where T : class {
            try {
                IPowersheetImporter<T> importer = PowersheetImportFactory.Get<T>(@"Files/sheet.xlsx");
                object[,] grid = importer.ToGrid(tableId);

                int y = (grid.GetLength(0) < 10) ? grid.GetLength(0) : 10;
                int x = (grid.GetLength(1) < 10) ? grid.GetLength(1) : 10;

                for (int o = 0; o < y; o++) {
                    for (int i = 0; i < x; i++) {
                        System.Console.Write(grid[o, i]);
                    }
                }
            }
            catch (Exception ex) {
                System.Console.WriteLine("Oops! {0}", ex.Message);
            }
        }

        static void RunImport<T>(int tableId) where T : class {
            try {
                IPowersheetImporter<T> importer = PowersheetImportFactory.Get<T>(@"Files/sheet.xlsx");
                List<T> data = importer.GetAll(tableId).ToList();

                foreach (var d in data) {
                    System.Console.WriteLine(d.ToString());
                    System.Console.WriteLine();
                }
            }
            catch (Exception ex) {
                System.Console.WriteLine("Oops! {0}", ex.Message);
            }
        }

        static void RunImport<T>(int tableId, int headingsRow, params string[] columns) where T : class {
            try {
                IPowersheetImporter<T> importer = PowersheetImportFactory.Get<T>(@"Files/sheet.xlsx");

                IPowersheetPropertyMap[] mappings = importer.GetMappings(columns).ToArray();
                List<T> data = importer.Fetch(tableId, headingsRow, null, null, mappings).ToList();

                foreach (var d in data) {
                    System.Console.WriteLine(d.ToString());
                    System.Console.WriteLine();
                }
            }
            catch (Exception ex) {
                System.Console.WriteLine("Oops! {0}", ex.Message);
            }
        }

        static void RunDump<T>(int tableId) {
            try {
                IPowersheetImporter<T> importer = PowersheetImportFactory.Get<T>(@"Files/sheet.xlsx");
                List<T> data = importer.GetAll(tableId).ToList();

                var dumpStuff = new List<DumpClass>();
                foreach (var d in data) {
                    var c = new DumpClass();
                }
            }
            catch (Exception ex) {
                System.Console.WriteLine("Oops! {0}", ex.Message);
            }
        }

        /// <summary>
        /// Simply for testing...
        /// </summary>
        static void TestDump() {
            IPowersheetImporter<Movie> importer = PowersheetImportFactory.Get<Movie>(@"Files/sheet.xlsx");
            List<Movie> data = importer.GetAll(0).ToList();

            foreach (var d in data) {
                d.Columns.Add("test", "this is dumped!");
            }

            IPowersheetExporter exporter = PowersheetExportFactory.Get();
            StringBuilder dump = exporter.Dump(data, true, false);

            using (StreamWriter writer = new StreamWriter(@"dump.csv")) {
                writer.WriteLine(dump.ToString());
            }
        }
    }

    class DumpClass : IPowersheetDump {

        public Dictionary<string, string> Columns { get; set; }

        public DumpClass() {
            Columns = new Dictionary<string, string>();
        }
    }
}
