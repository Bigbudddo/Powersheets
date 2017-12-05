using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

using Excel;

namespace Powersheets {

    public static class PowersheetImportFactory {

        public static IPowersheetImporter<T> Get<T>(string filePath) {
            try {
                if (String.IsNullOrWhiteSpace(filePath)) {
                    throw new Exception("File path cannot be null!");
                }
                if (!File.Exists(filePath)) {
                    throw new Exception("File does not exist!");
                }

                string fileExt = filePath.ExtractFileExtension();
                using (FileStream stream = File.Open(filePath, FileMode.Open)) {
                    return GetImporter<T>(stream, fileExt);
                }
            }
            catch (Exception ex) {
                throw new Exception(ex.Message);
            }
        }

        public static IPowersheetImporter<T> Get<T>(HttpRequestBase httpRequest, int fileIndex = 0) {
            try {
                if (httpRequest.Files.Count == 0 || fileIndex >= httpRequest.Files.Count) {
                    throw new Exception("No File available!");
                }

                HttpPostedFileBase file = httpRequest.Files[fileIndex];
                string fileExt = file.FileName.ExtractFileExtension();
                return GetImporter<T>(file.InputStream, fileExt);
            }
            catch (Exception ex) {
                throw new Exception(ex.Message);
            }
        }

        public static IPowersheetImporter Get(string filePath) {
            try {
                if (String.IsNullOrWhiteSpace(filePath)) {
                    throw new Exception("File path cannot be null!");
                }
                if (!File.Exists(filePath)) {
                    throw new Exception("File does not exist!");
                }

                string fileExt = filePath.ExtractFileExtension();
                using (FileStream stream = File.Open(filePath, FileMode.Open)) {
                    return GetImporter(stream, fileExt);
                }
            }
            catch (Exception ex) {
                throw new Exception(ex.Message);
            }
        }

        public static IPowersheetImporter Get(HttpRequestBase httpRequest, int fileIndex = 0) {
            try {
                if (httpRequest.Files.Count == 0 || fileIndex >= httpRequest.Files.Count) {
                    throw new Exception("No File available!");
                }

                HttpPostedFileBase file = httpRequest.Files[fileIndex];
                string fileExt = file.FileName.ExtractFileExtension();
                return GetImporter(file.InputStream, fileExt);
            }
            catch (Exception ex) {
                throw new Exception(ex.Message);
            }
        }

        private static IPowersheetImporter<T> GetImporter<T>(Stream stream, string fileExt) {
            IExcelDataReader reader;
            switch (fileExt.ToUpper()) {
                case "XLS":
                    reader = ExcelReaderFactory.CreateBinaryReader(stream);
                    break;
                case "XLSX":
                    reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    break;
                default:
                    throw new Exception("Un-recognised File Extension!");
            }

            if (reader != null && !String.IsNullOrWhiteSpace(fileExt)) {
                switch (fileExt.ToUpper()) {
                    case "XLS":
                    case "XLSX":
                        DataSet data = reader.AsDataSet();
                        return new SpreadsheetImporter<T>(data);
                    default:
                        throw new Exception("Un-recognised File Extension!");
                }
            }
            else {
                throw new Exception("Unable to read file!");
            }
        }

        private static IPowersheetImporter GetImporter(Stream stream, string fileExt) {
            IExcelDataReader reader;
            switch (fileExt.ToUpper()) {
                case "XLS":
                    reader = ExcelReaderFactory.CreateBinaryReader(stream);
                    break;
                case "XLSX":
                    reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                    break;
                default:
                    throw new Exception("Un-recognised File Extension!");
            }

            if (reader != null && !String.IsNullOrWhiteSpace(fileExt)) {
                switch (fileExt.ToUpper()) {
                    case "XLS":
                    case "XLSX":
                        DataSet data = reader.AsDataSet();
                        return new SpreadsheetImporter(data);
                    default:
                        throw new Exception("Un-recognised File Extension!");
                }
            }
            else {
                throw new Exception("Unable to read file!");
            }
        }
    }
}
