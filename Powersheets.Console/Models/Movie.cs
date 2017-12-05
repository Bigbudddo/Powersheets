using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Powersheets;

namespace Powersheets.Console {

    [SpreadsheetAutoIncrement("Id")]
    public class Movie : IPowersheetDump {

        public int Id { get; set; }

        [SpreadsheetColumn("Film Name", "Name")]
        public string Name { get; set; }

        [SpreadsheetColumn("Film Director", "Director")]
        public string Director { get; set; }

        [SpreadsheetColumn("Film Release Date", "FilmReleaseDate")]
        public DateTime FilmReleaseDate { get; set; }

        [SpreadsheetColumn("Age", "AgeRating")]
        public int AgeRating { get; set; }

        [SpreadsheetColumn("Price")]
        public decimal Price { get; set; }

        public Dictionary<string, string> Columns { get; set; }

        public Movie() {
            Columns = new Dictionary<string, string>();
        }
    }
}
