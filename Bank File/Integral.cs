using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace Bank_File
{
    public class Integral
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
                    
                    int adddetailsCol= Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Add Details 5");
                    int remarksCol = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "remark");
                    int dateCol = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Date");
                    int beneficiaryNameCol = Program.getColumnNumber(filePath, inputWorkSheet.ToString(), "Beneficiary Name");
                    using (var outputPackage = new ExcelPackage())
                    {
                        var outputWorksheet = outputPackage.Workbook.Worksheets.Add(inputWorkSheet.ToString());
                        var sourceRange = inputWorkSheet.Cells[1, 1, lastRow, lastCol];
                        var destinationRange = outputWorksheet.Cells[1, 1, lastRow, lastCol];
                        sourceRange.Copy(destinationRange);
                        for (int row = 2; row <= lastRow; row++)
                        {
                            if (Program.ShrinkString(outputWorksheet.Cells[row, dateCol].Text) != "") { 
                            string dateText = outputWorksheet.Cells[row, dateCol].Text;
                            if (DateTime.TryParseExact(dateText, "dd-MM-yyyy",System.Globalization.CultureInfo.InvariantCulture,
                                System.Globalization.DateTimeStyles.None, out DateTime parsedDate))
                            {
                                outputWorksheet.Cells[row, adddetailsCol].Value = parsedDate.ToString("MMMM yyyy");
                            }
                            else
                            {
                                outputWorksheet.Cells[row, adddetailsCol].Value = ""; // or handle invalid format
                            }
                            outputWorksheet.Cells[row, adddetailsCol].Value = outputWorksheet.Cells[row, adddetailsCol].Text + " " + outputWorksheet.Cells[row, beneficiaryNameCol].Text;
                            outputWorksheet.Cells[row, remarksCol].Value = "IAS India " + parsedDate.ToString("MMMM yyyy") + " Salary";
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
