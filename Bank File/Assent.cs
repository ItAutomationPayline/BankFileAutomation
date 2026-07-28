using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace Bank_File
{
    public class Assent
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
                    //int transactionCol = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Transaction Reference");
                    using (var outputPackage = new ExcelPackage())
                    {
                        var outputWorksheet = outputPackage.Workbook.Worksheets.Add(inputWorkSheet.ToString());
                        var sourceRange = inputWorkSheet.Cells[1, 1, lastRow, lastCol];
                        var destinationRange = outputWorksheet.Cells[1, 1, lastRow, lastCol];
                        sourceRange.Copy(destinationRange);
                        int RecordTypeCol = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Record Type");
                        int amountCol = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Payment amount");
                        int paymentTypeCol = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Payment Type");
                        int processingModeCol = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Processing Mode");
                        int totalCount = 0;
                        double totalAmount = 0;
                        outputWorksheet.InsertRow(2, 1);
                        outputWorksheet.Cells[2, RecordTypeCol].Value = "H";
                        outputWorksheet.Cells[2, paymentTypeCol].Value = "P";
                        lastRow = outputWorksheet.Dimension.End.Row;
                        for (int row = 2; row <= lastRow; row++)
                        {
                            if (Program.ShrinkString(outputWorksheet.Cells[row, RecordTypeCol].Text) == "p") 
                            {
                                totalCount++;
                                totalAmount += outputWorksheet.Cells[row, amountCol].GetValue<double>();
                            }
                            if (Program.ShrinkString(outputWorksheet.Cells[row, RecordTypeCol].Text) == "")
                            {
                                outputWorksheet.Cells[row, RecordTypeCol].Value = "T";
                                outputWorksheet.Cells[row, paymentTypeCol].Value = totalCount;
                                outputWorksheet.Cells[row, processingModeCol].Value = totalAmount;
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