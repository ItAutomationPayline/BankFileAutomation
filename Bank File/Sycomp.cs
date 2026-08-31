using OfficeOpenXml;
using System;
using System.IO;

namespace Bank_File
{
    public class Sycomp
    {
        public static void BankFile_Automation(string filePath, string outputFilePath)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // Template path
                string today = DateTime.Today.ToString("yyyyMMdd");
                string templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                                                   "Config",
                                                   "Sycomp",
                                                   $"BLKPAY_YYYYMMDD.xlsx");

                using (var inputPackage = new ExcelPackage(new FileInfo(filePath)))
                using (var templatePackage = new ExcelPackage(new FileInfo(templatePath)))
                using (var outputPackage = new ExcelPackage())
                {
                    var inputWorksheet = inputPackage.Workbook.Worksheets[0];
                    var templateWorksheet = templatePackage.Workbook.Worksheets[0];

                    int lastRow = inputWorksheet.Dimension.End.Row;
                    int lastCol = inputWorksheet.Dimension.End.Column;

                    int ifscCol = Program.getColumnNumber(filePath, inputWorksheet.Name, "IFSC");
                    int transactionTypeCol = Program.getColumnNumber(filePath, inputWorksheet.Name, "Transaction Type");

                    var outputWorksheet = outputPackage.Workbook.Worksheets.Add(inputWorksheet.Name);

                    // Copy template header (A1:O2)
                    templateWorksheet.Cells["A1:O2"].Copy(outputWorksheet.Cells["A1"]);

                    // Copy input data starting from row 3
                    inputWorksheet.Cells[1, 1, lastRow, lastCol]
                                  .Copy(outputWorksheet.Cells[3, 1]);

                    // Since data starts at row 3, adjust row numbers
                    for (int row = 4; row <= lastRow + 2; row++)
                    {
                        string ifsc = Program.ShrinkString(outputWorksheet.Cells[row, ifscCol].Text);

                        if (ifsc.Contains("idfb"))
                        {
                            outputWorksheet.Cells[row, transactionTypeCol].Value = "IFT";
                            outputWorksheet.Cells[row, ifscCol].Value = "";
                        }
                        else if (!string.IsNullOrWhiteSpace(ifsc))
                        {
                            outputWorksheet.Cells[row, transactionTypeCol].Value = "NEFT";
                        }
                        if (outputWorksheet.Cells[row, 2].Text.Replace(" ","")!="") 
                        { outputWorksheet.Cells[row, 10].Value = "Sycomp – Salary – " + DateTime.Now.ToString("MMMM yyyy"); }
                    }

                    outputWorksheet.Cells.AutoFitColumns();
                    outputWorksheet.DeleteRow(3);
                    // Output filename: BLKPAY_YYYYMMDD.xlsx
                    string outputFile = Path.Combine(outputFilePath, $"BLKPAY_{today}.xlsx");

                    outputPackage.SaveAs(new FileInfo(outputFile));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}