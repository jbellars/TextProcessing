using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace FormatTextToSpeech
{
    /*
     * Purpose: to fix a partially-corrupted scanned document of typical OCR typos, then make substitutions of numbers so that a TTS (text-to-speech) program or the "Read Aloud" function on a Kindle can read years like English-speaking humans would "nineteen-eighty", rather than as numbers like 1,980 or "one thousand eighty".
     * Author: Justin Bellars
     * Created: June 2014
     * 
     * Plans: read in word lists from Excel Spreadsheet to create a substitution list.
     */
    class Program
    {
        // XSLX - Excel 2007, 2010, 2012, 2013
        private const string strProvider = "Microsoft.ACE.OLEDB.12.0";
        private const string strExtendedProps = "Excel 12.0 XML";
        private const string strDataSource = @"C:\cygwin64\home\jbellars\Dictionary.xlsx";

        static void Main(string[] args)
        {
            var dctReplacements = LoadDictionaryFromExcelFile(); //new Dictionary<string, string>();

            var inputFile = File.ReadAllText(@"C:\cygwin64\home\jbellars\ogbcp-orig.txt");

            inputFile = dctReplacements.Aggregate(inputFile, (current, entry) => Regex.Replace(current, entry.Key, entry.Value));

            File.WriteAllText(@"C:\cygwin64\home\jbellars\ogbcp-test.txt", inputFile);

            #region original code

            // Fix corrupted characters
            inputFile = FixCorruptedCharactersInTheInputFile(inputFile);

            // Fix references
            inputFile = FixReferencesInTheInputFile(inputFile);

            // Make computer-readable decades
            inputFile = MakeComputerReadableDecades(inputFile);

            // Make computer-readable years
            inputFile = MakeComputerReadableYears(inputFile);

            inputFile = Regex.Replace(inputFile, @"\b(\d{2})[s]", "");

            // Fix blatant errors in the text
            inputFile = FixTyposInText(inputFile);

            inputFile = SubstituteWordsForAbbreviations(inputFile);

            //MatchCollection matches = Regex.Matches(inputFile, @"\do");

            File.WriteAllText(@"C:\cygwin64\home\jbellars\ogbcp-output.txt", inputFile);

            //foreach (Match match in matches)
            //{
            //    foreach (Capture capture in match.Captures)
            //    {
            //        Console.WriteLine("Index={0}, Value={1}", match.Index, capture.Value);
            //    }
            //}
            #endregion original code

            Console.WriteLine("Done.");
            Console.ReadKey();
        }

        private KeyValuePair<string, string> GetKeyValuePairFromSpreadsheet()
        {
            throw new NotImplementedException();
            //using (SpreadsheetDocument inputXslxFile = SpreadsheetDocument.Open(strDataSource, false))
            //{
            //    var wb = inputXslxFile.WorkbookPart.Workbook;
            //    var workSheets = wb.Descendants<Sheet>();
            //    var sharedStrings = workSheets.GetFirstChild();

            //}
        }

        private static string GetConnectionString()
        {
            Dictionary<string, string> props = new Dictionary<string, string>();
            props["Provider"] = strProvider;
            props["Extended Properties"] = strExtendedProps;
            props["Data Source"] = strDataSource;

            StringBuilder sb = new StringBuilder();

            foreach (KeyValuePair<string, string> prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }



        private static Dictionary<string, string> LoadDictionaryFromExcelFile()
        {
            throw new NotImplementedException();
            //string connectionString = GetConnectionString();

            //var workbook = 

            //return substitutions;
        }


        #region Methods

        private static string MakeComputerReadableDecades(string inputFile)
        {
            // Decades
            inputFile = Regex.Replace(inputFile, @"\b1530s\b", "Fifteen-thirties");
            inputFile = Regex.Replace(inputFile, @"\b1540s\b", "Fifteen-fourties");
            inputFile = Regex.Replace(inputFile, @"\b1580s\b", "Fifteen-eighties");
            
            inputFile = Regex.Replace(inputFile, @"\b1630s\b", "Sixteen-thirties");
            inputFile = Regex.Replace(inputFile, @"\b1640s\b", "Sixteen-fourties");
            inputFile = Regex.Replace(inputFile, @"\b1650s\b", "Sixteen-fifties");
            inputFile = Regex.Replace(inputFile, @"\b1660s\b", "Sixteen-sixties");
            inputFile = Regex.Replace(inputFile, @"\b1680s\b", "Sixteen-eighties");
            
            inputFile = Regex.Replace(inputFile, @"\b1720s\b", "Seventeen-twenties");
            inputFile = Regex.Replace(inputFile, @"\b1730s\b", "Seventeen-thirties");
            inputFile = Regex.Replace(inputFile, @"\b1770s\b", "Seventeen-seventies");
            inputFile = Regex.Replace(inputFile, @"\b1790s\b", "Seventeen-nineties");

            inputFile = Regex.Replace(inputFile, @"\b1800s\b", "Eighteen-hundreds");
            inputFile = Regex.Replace(inputFile, @"\b1820s\b", "Eighteen-twenties");
            inputFile = Regex.Replace(inputFile, @"\b1830s\b", "Eighteen-thirties");
            inputFile = Regex.Replace(inputFile, @"\b1840s\b", "Eighteen-fourties");
            inputFile = Regex.Replace(inputFile, @"\b1850s\b", "Eighteen-fifties");
            inputFile = Regex.Replace(inputFile, @"\b1870s\b", "Eighteen-seventies");
            inputFile = Regex.Replace(inputFile, @"\b1890s\b", "Eighteen-nineties");
            
            inputFile = Regex.Replace(inputFile, @"\b1900s\b", "Nineteen-hundreds");
            inputFile = Regex.Replace(inputFile, @"\b1920s\b", "Nineteen-twenties");
            inputFile = Regex.Replace(inputFile, @"\b1930s\b", "Nineteen-thirties");
            inputFile = Regex.Replace(inputFile, @"\b1940s\b", "Nineteen-fourties");
            inputFile = Regex.Replace(inputFile, @"\b1950s\b", "Nineteen-fifties");
            inputFile = Regex.Replace(inputFile, @"\b1960s\b", "Nineteen-sixties");
            inputFile = Regex.Replace(inputFile, @"\b1970s\b", "Nineteen-seventies");
            inputFile = Regex.Replace(inputFile, @"\b1980s\b", "Nineteen-eighties");
            inputFile = Regex.Replace(inputFile, @"\b1990s\b", "Nineteen-nineties");
            return inputFile;
        }

        private static string MakeComputerReadableYears(string inputFile)
        {
            

            // Specific Years
            inputFile = Regex.Replace(inputFile, @"\b1085\b", "Ten-eighty-five");
            inputFile = Regex.Replace(inputFile, @"\b1152\b", "Eleven-fifty-two");
            inputFile = Regex.Replace(inputFile, @"\b1200\b", "Twelve-hundred");
            inputFile = Regex.Replace(inputFile, @"\b1215\b", "Twelve-fifteen");
            inputFile = Regex.Replace(inputFile, @"\b1250\b", "Twelve-fifty");
            inputFile = Regex.Replace(inputFile, @"\b1267\b", "Twelve-sixty-seven");
            inputFile = Regex.Replace(inputFile, @"\b1281\b", "Twelve-eighty-one");
            inputFile = Regex.Replace(inputFile, @"\b1350\b", "Thirteen-fifty");
            inputFile = Regex.Replace(inputFile, @"\b1400\b", "Fourteen-hundred");
            inputFile = Regex.Replace(inputFile, @"\b1455\b", "Fourteen-hundred-fifty-five");
            inputFile = Regex.Replace(inputFile, @"\b1472\b", "Fourteen-hundred-seventy-two");
            inputFile = Regex.Replace(inputFile, @"\b1480\b", "Fourteen-hundred-eighty"); 
            inputFile = Regex.Replace(inputFile, @"\b1489\b", "Fourteen-hundred-eighty-nine");
            inputFile = Regex.Replace(inputFile, @"\b1500\b", "Fifteen-hundred");
            inputFile = Regex.Replace(inputFile, @"\b1508\b", "Fifteen-o-eight");
            inputFile = Regex.Replace(inputFile, @"\b1509\b", "Fifteen-o-nine");
            inputFile = Regex.Replace(inputFile, @"\b1515\b", "Fifteen-fifteen");
            inputFile = Regex.Replace(inputFile, @"\b1517\b", "Fifteen-seventeen");
            inputFile = Regex.Replace(inputFile, @"\b1520\b", "Fifteen-twenty");
            inputFile = Regex.Replace(inputFile, @"\b1544\b", "Fifteen-fourty-four");
            inputFile = Regex.Replace(inputFile, @"\b1549\b", "Fifteen-fourty-nine");
            inputFile = Regex.Replace(inputFile, @"\b1552\b", "Fifteen-fifty-two");
            inputFile = Regex.Replace(inputFile, @"\b1556\b", "Fifteen-fifty-six");

            inputFile = Regex.Replace(inputFile, @"\b\b", "");

            inputFile = Regex.Replace(inputFile, @"\b1800\b", "Eighteen-hundred");

            inputFile = Regex.Replace(inputFile, @"\b1900\b", "Nineteen-hundred");
            return inputFile;
        }

        private static string FixReferencesInTheInputFile(string inputFile)
        {
            inputFile = Regex.Replace(inputFile, @"[Ii]{2}( Tim[\.othy]?)", @"2${1}");
            inputFile = Regex.Replace(inputFile, @"[Ii]( Tim[\.othy]?)", @"1${1}");
            inputFile = Regex.Replace(inputFile, @"[Ii]{3}( John)", @"3${1}");
            inputFile = Regex.Replace(inputFile, @"[Ii]{2}( John)", @"2${1}");
            inputFile = Regex.Replace(inputFile, @"[Ii]( John)", @"1${1}");
            inputFile = Regex.Replace(inputFile, @"[Ii]{2}( Sam[\.uel]?)", @"2${1}");
            inputFile = Regex.Replace(inputFile, @"[Ii]( Sam[\.uel]?)", @"1${1}");
            return inputFile;
        }

        private static string FixCorruptedCharactersInTheInputFile(string inputFile)
        {
            inputFile = Regex.Replace(inputFile, @"[Ii][Oo](\d)", @"10${1}");
            inputFile = Regex.Replace(inputFile, @"[Oo][Ii](\d)", @"01${1}");
            inputFile = Regex.Replace(inputFile, @"(\d)[Oo][Ii]", @"${1}01");
            inputFile = Regex.Replace(inputFile, @"(\d)oo", @"${1}00");
            inputFile = Regex.Replace(inputFile, @"(\d)o", @"${1}0");
            inputFile = Regex.Replace(inputFile, @"o(\d)", @"0${1}");
            inputFile = Regex.Replace(inputFile, @"[g](\d)", @"9${1}");
            inputFile = Regex.Replace(inputFile, @"[Ii]{2}(\d)", @"11${1}");
            inputFile = Regex.Replace(inputFile, @"[Ii](\d)", @"1${1}");
            inputFile = Regex.Replace(inputFile, @"[Zz](\d)", @"1${1}");
            inputFile = Regex.Replace(inputFile, @"(\d)[Zz]", @"${1}1");
            inputFile = Regex.Replace(inputFile, @"(\d)[Ii]{2}", @"${1}11");
            inputFile = Regex.Replace(inputFile, @"(\d)[Ii]", @"${1}1");
            inputFile = Regex.Replace(inputFile, @"(\d{3})([0])[5]\b", @"${1}${2}s");
            inputFile = Regex.Replace(inputFile, @"(\d)(\d{3})\1(\d{3})", "${1}${2} - ${1}${3}");
            return inputFile;
        }

        private static string SubstituteWordsForAbbreviations(string inputFile)
        {
            inputFile = Regex.Replace(inputFile, @"\bDr\s", "Doctor");
            return inputFile;
        }

        private static string FixTyposInText(string inputFile)
        {
            // Fix typos in the text
            inputFile = Regex.Replace(inputFile, @"\bFrankfurt-am-Oder\b", "Frankfurt-an-der-Oder");

            return inputFile;
        }

        #endregion Methods
    }
}
