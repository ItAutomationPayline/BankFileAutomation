using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;

namespace Bank_File
{
    public class Synchronoss
    {
        public static void BankFile_Automation(string filePath,string outputFilePath)
        {
            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {
                    int IP = Program.getSheetNumber(filePath, "RepGenSal");
                    var inputWorkSheet = package.Workbook.Worksheets[IP]; // Get the input worksheet
                    int lastRow = inputWorkSheet.Dimension.End.Row;
                    int lastCol = inputWorkSheet.Dimension.End.Column;

                    using (var outputPackage = new ExcelPackage())
                    {
                        var outputRepGenWorksheet = outputPackage.Workbook.Worksheets.Add("RepGenSal");
                        var outputHDFCsheet = outputPackage.Workbook.Worksheets.Add("HDFC");
                        var outputOtherBankSheet = outputPackage.Workbook.Worksheets.Add("Other Bank");
                        var outputFnFSheet = outputPackage.Workbook.Worksheets.Add("F&F");
                        //inputWorkSheet.Cells[lastRow, 9].Formula = "SUM(I:I)";
                        //// Copy all data, including values, formulas, and formatting
                        var sourceRange = inputWorkSheet.Cells[1, 1, lastRow, lastCol];
                        var destinationRange = outputRepGenWorksheet.Cells[1, 1, lastRow, lastCol];
                        sourceRange.Copy(destinationRange);
                        destinationRange = outputHDFCsheet.Cells[1, 1, lastRow, lastCol];
                        sourceRange.Copy(destinationRange);
                        destinationRange = outputOtherBankSheet.Cells[1, 1, lastRow, lastCol];
                        sourceRange.Copy(destinationRange);
                        destinationRange = outputFnFSheet.Cells[1, 1, lastRow, lastCol];
                        sourceRange.Copy(destinationRange);
                        int hdfccol = Program.getColumnNumber(filePath, "RepGenSal", "Primary Bank Name");
                        int leftdatecol = Program.getColumnNumber(filePath, "RepGenSal", "Left Date");
                        for(int row= lastRow - 2;row>=3;row--)
                        {
                            if (!outputHDFCsheet.Cells[row,hdfccol].Text.Contains("HDFC") || outputHDFCsheet.Cells[row, leftdatecol].Text != "")
                            {
                                outputHDFCsheet.DeleteRow(row);
                            }
                            if (outputOtherBankSheet.Cells[row, hdfccol].Text.Contains("HDFC")|| outputOtherBankSheet.Cells[row, leftdatecol].Text != "")
                            {
                                outputOtherBankSheet.DeleteRow(row);
                            }
                            if (outputFnFSheet.Cells[row, leftdatecol].Text=="")
                            {
                                outputFnFSheet.DeleteRow(row);
                            }
                        }
                        lastRow = outputOtherBankSheet.Dimension.End.Row;
                        int prevrow = lastRow - 1;
                        int prevprevrow = lastRow - 2;
                        outputOtherBankSheet.Cells[prevrow, 9].Formula = "SUM(I3:I" + prevprevrow + ")";
                        
                        double outputotherbanksum = outputOtherBankSheet.Cells[prevrow, 9].GetValue<double>();
                        for (int row=3;row<=prevrow;row++)
                        {
                            outputotherbanksum = outputotherbanksum + outputOtherBankSheet.Cells[row, 9].GetValue<double>();
                        }
                        lastRow = outputHDFCsheet.Dimension.End.Row;
                        prevrow = lastRow - 1;
                        prevprevrow = lastRow - 2;
                        outputHDFCsheet.Cells[prevrow, 9].Formula = "SUM(I3:I" + prevprevrow + ")";
                        double outputhdfcsum = outputHDFCsheet.Cells[prevrow, 9].GetValue<double>();
                        for (int row = 3; row <= prevrow; row++)
                        {
                            outputhdfcsum = outputhdfcsum + outputHDFCsheet.Cells[row, 9].GetValue<double>();
                        }
                        lastRow = outputFnFSheet.Dimension.End.Row;
                        prevrow = lastRow - 1;
                        prevprevrow = lastRow - 2;
                        outputFnFSheet.Cells[prevrow, 9].Formula = "SUM(I3:I" + prevprevrow + ")";
                        double outputfnfsum = outputFnFSheet.Cells[prevrow, 9].GetValue<double>();
                        for (int row = 3; row <= prevrow; row++)
                        {
                            outputfnfsum = outputfnfsum + outputFnFSheet.Cells[row, 9].GetValue<double>();
                        }
                        lastRow = outputRepGenWorksheet.Dimension.End.Row;
                        prevrow = lastRow - 1;
                        prevprevrow = lastRow - 2;
                        double repgentotal = outputRepGenWorksheet.Cells[prevrow, 9].GetValue<double>();
                        //outputRepGenWorksheet.Cells[prevrow, 9].Formula = "SUM(I3:I" + prevprevrow + ")";
                        outputRepGenWorksheet.Cells[lastRow + 2, 8].Value = "HDFC";
                        outputRepGenWorksheet.Cells[lastRow + 2, 9].Value = outputhdfcsum;
                        outputRepGenWorksheet.Cells[lastRow + 3, 8].Value = "Other Bank";
                        outputRepGenWorksheet.Cells[lastRow + 3, 9].Value = outputotherbanksum;
                        outputRepGenWorksheet.Cells[lastRow + 4, 8].Value = "F&F";
                        outputRepGenWorksheet.Cells[lastRow + 4, 9].Value = outputfnfsum;
                        outputRepGenWorksheet.Cells[lastRow + 5, 8].Value = "Total";
                        outputRepGenWorksheet.Cells[lastRow + 5, 9].Value = outputhdfcsum+ outputotherbanksum+ outputfnfsum;
                        outputRepGenWorksheet.Cells[lastRow + 7, 8].Value = "As Per Register";
                        outputRepGenWorksheet.Cells[lastRow + 7, 9].Value = repgentotal;
                        outputRepGenWorksheet.Cells[lastRow + 8, 8].Value = "Difference";
                        outputRepGenWorksheet.Cells[lastRow + 8, 9].Value = repgentotal-(outputhdfcsum + outputotherbanksum + outputfnfsum);
                        // Optional: save the output package to a file
                        string newFileName = Path.Combine(outputFilePath, "Automated Bank File" + Path.GetFileName(filePath));
                        FileInfo newFileInfo = new FileInfo(newFileName);
                        outputRepGenWorksheet.Cells[outputRepGenWorksheet.Dimension.Address].AutoFitColumns();
                        outputHDFCsheet.Cells[outputHDFCsheet.Dimension.Address].AutoFitColumns();
                        outputOtherBankSheet.Cells[outputOtherBankSheet.Dimension.Address].AutoFitColumns();
                        outputFnFSheet.Cells[outputFnFSheet.Dimension.Address].AutoFitColumns();
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

