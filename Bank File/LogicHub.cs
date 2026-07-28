using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace Bank_File
{
    public class LogicHub
    {
        public static void BankFile_Automation(string filePath, string outputFilePath)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    var inputWorkSheet = package.Workbook.Worksheets[0]; // Get the input worksheet
                    int lastRow = inputWorkSheet.Dimension.End.Row;
                    int lastCol = inputWorkSheet.Dimension.End.Column;
                    int transactionTypeCol = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Transaction Type");
                    int beneBankNameCol = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Bene Bank Name");
                    using (var outputPackage = new ExcelPackage())
                    {
                        var outputWorksheet = outputPackage.Workbook.Worksheets.Add(inputWorkSheet.ToString());
                        var sourceRange = inputWorkSheet.Cells[1, 1, lastRow, lastCol];
                        var destinationRange = outputWorksheet.Cells[1, 1, lastRow, lastCol];
                        sourceRange.Copy(destinationRange);
                        for (int row=2;row<=lastRow;row++)
                        {
                            if (Program.ShrinkString(outputWorksheet.Cells[row, beneBankNameCol].Text).Contains("hdfc"))
                            {
                                outputWorksheet.Cells[row, transactionTypeCol].Value = "N";
                            }
                            else if (!Program.ShrinkString(outputWorksheet.Cells[row, beneBankNameCol].Text).Contains("hdfc") && Program.ShrinkString(outputWorksheet.Cells[row, beneBankNameCol].Text)!="")
                            {
                                outputWorksheet.Cells[row, transactionTypeCol].Value = "I";
                            }
                        }
                        outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                        string newFileName = Path.Combine(outputFilePath, "Automated Bank File " + Path.GetFileName(filePath));
                        FileInfo newFileInfo = new FileInfo(newFileName);
                        outputPackage.SaveAs(newFileInfo);
                        outputPackage.SaveAsAsync(new FileInfo(outputFilePath));
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }

        }
    }
}