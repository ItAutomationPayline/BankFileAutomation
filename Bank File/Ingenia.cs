using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace Bank_File
{
    public class Ingenia
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
                    int hridCol= Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Employee Number");
                    int transactionCol = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Reference Number");
                    int batchrefno= Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Batch Ref no");
                    using (var outputPackage = new ExcelPackage())
                    {
                        var outputWorksheet = outputPackage.Workbook.Worksheets.Add("Bank File");
                        var sourceRange = inputWorkSheet.Cells[1, 1, lastRow, lastCol];
                        var destinationRange = outputWorksheet.Cells[1, 1, lastRow, lastCol];
                        sourceRange.Copy(destinationRange);

                        for (int row = 2; row <= lastRow; row++)
                        {
                            string batchref = (outputWorksheet.Cells[row, batchrefno].Text).Replace("SALARY ","SAL").Replace(" ","");
                            if (Program.ShrinkString(outputWorksheet.Cells[row, 1].Text) != "" || Program.ShrinkString(outputWorksheet.Cells[row, 3].Text) != "" || Program.ShrinkString(outputWorksheet.Cells[row, 5].Text) != "")
                            {
                                outputWorksheet.Cells[row, transactionCol].Value = Program.GenerateIngenia15DigitUTR(outputWorksheet.Cells[row, hridCol].Text);
                                outputWorksheet.Cells[row, batchrefno].Value= Program.GenerateIngenia15DigitBatchRefNo(batchref);
                            }
                        }
                        outputWorksheet.Cells[outputWorksheet.Dimension.Address].AutoFitColumns();
                        outputWorksheet.DeleteColumn(hridCol);
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
