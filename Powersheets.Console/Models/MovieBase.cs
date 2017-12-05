using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Powersheets;

namespace Powersheets.Console {
    
    internal class MovieBase {

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

        public override string ToString() {
            return String.Format("{0} directed by {1} was released on {2} with an age rating of {3} and a ticket sale price of {4}",
                FilmName, Director, ReleaseDate.ToShortDateString(), AgeRating, TicketPrice);
        }
    }

    internal class MovieBaseMath {

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
        [SpreadsheetColumn("=MATH(ADD(Price,Glasses Cost))")]
        public decimal FinalTicketPrice { get; set; }
        [SpreadsheetColumn("=MATH(SUBTRACT(Price,Promotional Discount))")]
        public decimal PromotionalDiscount { get; set; }

        public override string ToString() {
            return String.Format("{0} directed by {1} was released on {2} with an age rating of {3} and a ticket sale price of {4} ({5} with Glasses & Promo Discount of {6})",
                FilmName, Director, ReleaseDate.ToShortDateString(), AgeRating, TicketPrice, FinalTicketPrice, PromotionalDiscount);
        }
    }

    internal class MovieBaseHeadings {

        [SpreadsheetColumn("=HEADING(LIKE(Film)7)")]
        public string FilmName { get; set; }
        [SpreadsheetColumn("=HEADING(EQUALS(Film Director)1)")]
        public string Director { get; set; }
        [SpreadsheetColumn("=HEADING(EQUALS(Film Release Date)1)")]
        public DateTime ReleaseDate { get; set; }
        [SpreadsheetColumn("Age")]
        public int AgeRating { get; set; }
        [SpreadsheetColumn("Price")]
        public decimal TicketPrice { get; set; }

        public override string ToString() {
            return String.Format("{0} directed by {1} was released on {2} with an age rating of {3} and a ticket sale price of {4}",
                FilmName, Director, ReleaseDate.ToShortDateString(), AgeRating, TicketPrice);
        }
    }

    internal class MovieBaseSet {
        [SpreadsheetColumn("Film Name")]
        public string FilmName { get; set; }
        [SpreadsheetColumn("Film Director")]
        public string Director { get; set; }
        [SpreadsheetColumn("=VALUE(SET(Film))")]
        public string Category { get; set; }
        [SpreadsheetColumn("Film Release Date")]
        public DateTime ReleaseDate { get; set; }
        [SpreadsheetColumn("Age")]
        public int AgeRating { get; set; }
        [SpreadsheetColumn("Price")]
        public decimal TicketPrice { get; set; }

        public override string ToString() {
            return String.Format("{0} directed by {1} was released on {2} with an age rating of {3} and a ticket sale price of {4} in category {5}",
                FilmName, Director, ReleaseDate.ToShortDateString(), AgeRating, TicketPrice, Category);
        }
    }

    internal class MovieBaseCast {
        [SpreadsheetColumn("Film Name")]
        public string FilmName { get; set; }
        [SpreadsheetColumn("Film Director")]
        public string Director { get; set; }
        [SpreadsheetColumn("=VALUE(SET(Film))")]
        public string Category { get; set; }
        [SpreadsheetColumn("=CAST(DATE(Film Release Date))")]
        public DateTime ReleaseDate { get; set; }
        [SpreadsheetColumn("Age")]
        public int AgeRating { get; set; }
        [SpreadsheetColumn("Price")]
        public decimal TicketPrice { get; set; }

        public override string ToString() {
            return String.Format("{0} directed by {1} was released on {2} with an age rating of {3} and a ticket sale price of {4} in category {5}",
                FilmName, Director, ReleaseDate.ToShortDateString(), AgeRating, TicketPrice, Category);
        }
    }

    internal class MovieBaseCategory {

        [SpreadsheetColumn("Film Name")]
        public string FilmName { get; set; }
        [SpreadsheetColumn("Film Director")]
        public string Director { get; set; }
        [SpreadsheetColumn("=SELECT(CELL(1, 0))")]
        public string Category { get; set; }
        [SpreadsheetColumn("=SELECT(LOOK(Category:,2,0))")]
        public string SubCategory { get; set; }
        [SpreadsheetColumn("=SELECT(WORKBOOK(Cinemas,0,1))")]
        public string Cinema { get; set; }
        [SpreadsheetColumn("=SELECT(WORKLOOK(Cinemas,Boxville,1,0))")]
        public string CinemaAddress { get; set; }
        [SpreadsheetColumn("Film Release Date")]
        public DateTime ReleaseDate { get; set; }
        [SpreadsheetColumn("Age")]
        public int AgeRating { get; set; }
        [SpreadsheetColumn("Price")]
        public decimal TicketPrice { get; set; }

        public override string ToString() {
            return String.Format("{0} in category {1}({7}) was directed by {2} was released on {3} with an age rating of {4} and a ticket sale price of {5} at the {6} Cinema ({8})",
                FilmName, Category, Director, ReleaseDate.ToShortDateString(), AgeRating, TicketPrice, Cinema, SubCategory, CinemaAddress);
        }
    }

    internal class MovieBaseCategoryResult {

        [SpreadsheetColumn("Film Name")]
        public string FilmName { get; set; }
        [SpreadsheetColumn("Film Director")]
        public string Director { get; set; }
        [SpreadsheetColumn("=SELECT(CELL(1, 0))")]
        public string Category { get; set; }
        [SpreadsheetColumn("=SELECT(LOOK(Category:,2,0))")]
        public string SubCategory { get; set; }
        [SpreadsheetColumn("Cinema")]
        public string Cinema { get; set; }
        [SpreadsheetColumn("=RESULT(WORKLOOK(Cinemas,Cinema,1,0))")]
        public string CinemaAddress { get; set; }
        [SpreadsheetColumn("Film Release Date")]
        public DateTime ReleaseDate { get; set; }
        [SpreadsheetColumn("Age")]
        public int AgeRating { get; set; }
        [SpreadsheetColumn("Price")]
        public decimal TicketPrice { get; set; }

        public override string ToString() {
            return String.Format("{0} in category {1}({7}) was directed by {2} was released on {3} with an age rating of {4} and a ticket sale price of {5} at the {6} Cinema ({8})",
                FilmName, Category, Director, ReleaseDate.ToShortDateString(), AgeRating, TicketPrice, Cinema, SubCategory, CinemaAddress);
        }
    }
}
