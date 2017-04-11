using System;
using System.Collections.Generic;
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

                //var importer = new PowersheetImporter<T>(@"Files/sheet.xlsx");
                //IEnumerable<IPowersheetPropertyMap> mappings = importer.GetPropertyMappings(columns);
                //List<T> data = importer.GetAll(tableId, headingsRow, mappings.ToArray()).ToList();

                foreach (var d in data) {
                    System.Console.WriteLine(d.ToString());
                    System.Console.WriteLine();
                }
            }
            catch (Exception ex) {
                System.Console.WriteLine("Oops! {0}", ex.Message);
            }
        }
    }
}
