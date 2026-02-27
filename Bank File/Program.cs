using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Principal;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Bank_File
{
    public class Program
    {
        public static string basePath = AppDomain.CurrentDomain.BaseDirectory;
        public static string msg = "";
        public static string sourceFolder = basePath;
        public static string outputFilePath = basePath + "/Output/";
        public static int row2 = 2;
        public static string filePath1 = "";
        private static readonly object _lock = new object();
        private static int _sequence = 0;
        private static DateTime _lastSecond = DateTime.MinValue;
        private static readonly Random _random = new Random();
        private const string Base36 = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        private const string Letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        public static string GenerateIngenia15DigitUTR(string hrid)
        {
            if (string.IsNullOrWhiteSpace(hrid))
                throw new ArgumentException("HRID cannot be null or empty.");

            hrid = hrid.Trim().ToUpper();

            if (hrid.Length >= 15)
                throw new ArgumentException("HRID length must be less than 15.");

            lock (_lock)
            {
                DateTime now = DateTime.UtcNow;

                if (now.ToString("yyMMddHHmmss") == _lastSecond.ToString("yyMMddHHmmss"))
                    _sequence++;
                else
                {
                    _sequence = 0;
                    _lastSecond = now;
                }

                int remainingLength = 15 - hrid.Length;

                // Time-based uniqueness source
                string timePart = now.ToString("yyMMddHHmmss");

                // Sequence in base36
                string seqPart = ToBase36(_sequence);

                // Random alphanumeric buffer
                string randomPart = new string(
                    Enumerable.Range(0, 5)
                    .Select(_ => Base36[_random.Next(Base36.Length)])
                    .ToArray()
                );

                string combined = timePart + seqPart + randomPart;

                // Take required remaining length
                string suffix = combined.Substring(0, remainingLength);

                // Ensure at least one alphabet in suffix
                if (!suffix.Any(char.IsLetter))
                {
                    suffix = suffix.Substring(0, suffix.Length - 1) +
                             Letters[_random.Next(Letters.Length)];
                }

                return hrid + suffix;
            }
        }
        public static string GenerateUTR()
        {
            lock (_lock)
            {
                DateTime now = DateTime.UtcNow;

                if (now.Second == _lastSecond.Second &&
                    now.Minute == _lastSecond.Minute &&
                    now.Hour == _lastSecond.Hour &&
                    now.Day == _lastSecond.Day)
                {
                    _sequence++;
                }
                else
                {
                    _sequence = 0;
                    _lastSecond = now;
                }
                // Time part: 12 digits
                string timePart = now.ToString("yyMMddHHmmss");
                // Sequence part: 2 base36 chars (1296 per second)
                string seqPart = ToBase36(_sequence).PadLeft(2, '0');
                // FORCE at least one alphabet
                char alpha = Letters[_random.Next(Letters.Length)];
                // Total = 12 + 2 + 1 = 15 → add one more random base36
                char extra = Base36[_random.Next(Base36.Length)];

                return timePart + seqPart + alpha + extra; // 16 chars
            }
        }
        public static string Generate15DigitUTR()
        {
            lock (_lock)
            {
                DateTime now = DateTime.UtcNow;

                if (now.Second == _lastSecond.Second &&
                    now.Minute == _lastSecond.Minute &&
                    now.Hour == _lastSecond.Hour &&
                    now.Day == _lastSecond.Day)
                {
                    _sequence++;
                }
                else
                {
                    _sequence = 0;
                    _lastSecond = now;
                }
                // Time part: 12 digits
                string timePart = now.ToString("yyMMddHHmmss");
                // Sequence part: 2 base36 chars (1296 per second)
                string seqPart = ToBase36(_sequence).PadLeft(2, '0');
                // FORCE at least one alphabet
                char alpha = Letters[_random.Next(Letters.Length)];
                // Total = 12 + 2 + 1 = 15 → add one more random base36

                return timePart + seqPart + alpha; // 16 chars
            }
        }
        private static string ToBase36(int value)
        {
            if (value == 0) return "0";

            string result = "";
            while (value > 0)
            {
                result = Base36[value % 36] + result;
                value /= 36;
            }
            return result;
        }
        public static string ShrinkString(string input)
        {
            if (input != null)
            {
                input = input.ToLower();
                input = input.Replace(" ", "");
                return input;
            }
            return "";
        }
        public static int getColumnNumber(string filepath, string worksheetname, string columnname)
        {
            try
            {
                columnname = columnname.ToLower();
                columnname = columnname.Replace(" ", "");
                using (var package = new ExcelPackage(new FileInfo(filepath)))
                {
                    var inputWorkSheet = package.Workbook.Worksheets[getSheetNumber(filePath1, worksheetname)];
                    int col = 1;
                    int row = 1;
                    int totalColumns = inputWorkSheet.Dimension.End.Column;
                    int totalRows = inputWorkSheet.Dimension.End.Row;
                    for (row = 1; row <= totalRows; row++) 
                    {
                        for (col = 1; col <= totalColumns; col++)
                        {
                            string temp = inputWorkSheet.Cells[row, col].Text.ToLower();
                            temp = temp.Replace(" ", "");
                            if (columnname.Equals(temp))
                            {
                                return col;
                            }
                        }
                    }
                    col = 1111;
                    if (col == 1111)
                    {
                        Console.WriteLine(columnname + " column was not found in " + worksheetname + " of " + filepath + " file.");
                        //ErrorCount++;
                    }
                    return col;
                }
            }
            catch (Exception e)
            {
                Log(columnname + " column was not found in" + worksheetname + " of " + Path.GetFileName(filepath) + " file.");
                Console.WriteLine(e);
                Console.WriteLine(columnname + " column was not found in" + worksheetname + " of " + filepath + " file.");
                Console.ReadLine();
                throw;
            }
        }
        public static int getSheetNumber(string filepath, string worksheetname)
        {
            try
            {
                worksheetname = ShrinkString(worksheetname);
                using (var package = new ExcelPackage(new FileInfo(filepath)))
                {
                    int worksheetCount = package.Workbook.Worksheets.Count;
                    int i = 0;
                    for (i = worksheetCount - 1; i >= 0; i--)
                    {
                        string temp = package.Workbook.Worksheets[i].Name;
                        temp = ShrinkString(temp);
                        if (temp.Equals(worksheetname))
                        {
                            return i;
                        }
                    }
                    i = 0;
                    if (i == 0)
                    {
                        Log(worksheetname + " sheet was not found in " + Path.GetFileName(filepath));
                    }
                    return i;
                }
            }
            catch (Exception e)
            {
                Log(e.Message);
                Console.ReadLine();
                throw;
            }
        }
        public static void Main(string[] args)
        {
            Console.WriteLine("After pasting Bank file, Press Enter to start processing.");
            Console.ReadLine();

            var inputFile = Directory.GetFiles(sourceFolder + "Input", "*.xls*")
                .OrderByDescending(f => new FileInfo(f).CreationTime).ToList();

            if (inputFile.Count == 0)
            {
                Console.WriteLine("Required file not found. Ensure there is .xlsx file.");
                return;
            }
            filePath1 = filePath1 + inputFile.First();

            Console.WriteLine($"Using file: {filePath1}");
            if (filePath1.ToLower().Contains("synchronoss"))
            {
                Synchronoss.BankFile_Automation(filePath1, outputFilePath);
            }
            if (filePath1.ToLower().Contains("foreflight")|| filePath1.ToLower().Contains("jepp"))
            {
                ForeFlight.BankFile_Automation(filePath1, outputFilePath);
            }
            if (filePath1.ToLower().Contains("ingenia"))
            {
                Ingenia.BankFile_Automation(filePath1, outputFilePath);
            }
            Console.WriteLine("Processing completed. Output file generated.");
            Console.ReadLine();
        }
        public static void Log(string message)
        {
            try
            {
                DateTime today = DateTime.Today;
                string _logFilePath = outputFilePath + Path.GetFileName(filePath1) + today.ToString("dd/MMMM/yyyy") + "_BankFileAutomation.log";
                Directory.CreateDirectory(Path.GetDirectoryName(_logFilePath));
                File.AppendAllText(_logFilePath, $"{DateTime.Now}: {message}{Environment.NewLine}");
                msg = msg + message + "<br>";
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }
    }
}
