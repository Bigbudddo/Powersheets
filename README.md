# Powersheets
A simple set of functions that allows you to turn Excel/Spreadsheets into Objects through the use of Attributes, which is built upon the [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) library. Contains basic-to-advanced Formula controls that allow building C# Objects from multiple Workbooks.

**NOTE**
This will only work with the following Spreadsheet File extensions currently:
* XLS
* XLSX

It can also export your Objects into the following formats:
* CSV
* XLS

## Getting Started
To get started, simply clone/download the repo to your local machine and open with Visual Studio. Ensure that you Clean and then Build the Solution, which will output the DLLs required into the bin folder. Copy the following required DLLs into your project and reference them accordingly.

> * ICSharpCode.SharpZipLib.dll
> * Powesheets.dll

## How to use
Consider the following Object/Model. The Attributes have been assigned to a Spreadsheet Column heading accordingly.
```c#
class Movie {

    [SpreadsheetColumn("Film Name")]
    public string FilmName { get; set; }
    [SpreadsheetColumn("Film Director")]
    public string Director { get; set; }
    [SpreadsheetColumn("Film Release Date")]
    public DateTime ReleaseDate { get; set; }
    [SpreadsheetColumn("Age")]
    public int AgeRating { get; set; }
    [SpreadsheetColumn("Price")]
    public decimal TicketPrice { get; set; }
}
```

Now, here is a simple code solution to read a **xlsx** file in the root folder called **sheet**
```c#
int tableId = 0;

IPowersheetImporter<Movie> importer = PowersheetImportFactory.Get<Movie>(@"sheet.xlsx");
List<Movie> data = importer.GetAll(tableId).ToList();
```

Need to extract something more specific?
```c#
int tableId = 0;
int? headingsRow = 1;
int? start = null;
int? limit = null;
string[] propertyNames = new string[] { "FilmName", "Director", "ReleaseDate", "AgeRating", "TicketPrice" };

IPowersheetImporter<Movie> importer = PowersheetImportFactory.Get<Movie>(@"sheet.xlsx");
IPowersheetPropertyMap[] mappings = importer.GetMappings(propertyNames).ToArray();
List<Movie> data = importer.Fetch(tableId, headingsRow, start, limit, mappings).ToList();
```

Or maybe, you just want the entire worksheet as some form of grid?
```c#
int tableId = 0;

IPowersheetImporter<Movie> importer = PowersheetImportFactory.Get<Movie>(@"sheet.xlsx");
object[,] grid = importer.ToGrid(tableId);
```

The **new** importer factory implementation now contains a way for MVC/Web projects to handle Spreadsheet File Uploads into Objects.

## Motivation
This project was inspired by my recent involvement with a work project that required multiple different Spreadsheet files to be read into to create a collection of a specific Object. These Spreadsheets had columns in different orders and the machine running the code, did not have access to Microsoft Excel. Fortunately they all had headings that could be read. Whilst this is **NOT** the solution to the work problem, it did make an interesting project to continue and hopefully can be used to solve similar problems for other individuals/companies.

## Built With
* [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) - Essentially the Core

## Authors
* **Stuart Harrison** - [Bigbudddo](https://github.com/Bigbudddo)

See also the list of [contributors](https://github.com/Bigbudddo/Powersheets/graphs/contributors) who participated in this project.

## License
This project is licensed under the MIT License - see the [LICENSE.md](https://github.com/Bigbudddo/Powersheets/blob/master/LICENSE) file or visit [MIT](https://choosealicense.com/licenses/mit/) for more details.

## Acknowledgments
This project is heavily inspired by the [PetaPoco](https://github.com/CollaboratingPlatypus/PetaPoco) project in that properties of an object are mapped to columns in a database.
