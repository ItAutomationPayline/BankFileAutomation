using System;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;

namespace Bank_File
{
    public class GH_Induction
    {
        public static void BankFile_Automation(string sasa, string outputFilePath)
        {
            // Get SBI File
            var sbiFile = Directory.GetFiles(
                    Path.Combine(Program.sourceFolder, "Input"),
                    "*.xls*")
                .Where(f =>
                    f.ToLower().Contains("sbi") ||
                    f.ToLower().Contains("state bank"))
                .FirstOrDefault();

            // Get Other Bank File
            var otherbankFile = Directory.GetFiles(
                    Path.Combine(Program.sourceFolder, "Input"),
                    "*.xls*").Where(f => f.ToLower().Contains("other"))
                .FirstOrDefault();

            // Validation
            if (string.IsNullOrEmpty(sbiFile))
            {
                Console.WriteLine("Required SBI file not found. Make sure there is gh induction and sbi written in name");
                return;
            }

            if (string.IsNullOrEmpty(otherbankFile))
            {
                Console.WriteLine("Other bank file not found.make sure there is gh induction and other bank written in name");
                return;
            }

            try
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (var outputPackage = new ExcelPackage())
                {
                    // =========================
                    // COPY SBI FILE FIRST SHEET
                    // =========================
                    using (var sbiPackage = new ExcelPackage(new FileInfo(sbiFile)))
                    {
                        var inputWorksheet = sbiPackage.Workbook.Worksheets.FirstOrDefault();

                        if (inputWorksheet != null && inputWorksheet.Dimension != null)
                        {
                            int lastRow = inputWorksheet.Dimension.End.Row;
                            int lastCol = inputWorksheet.Dimension.End.Column;

                            // Create Sheet
                            var outputWorksheet =
                                outputPackage.Workbook.Worksheets.Add("State Bank");

                            // Copy Data
                            var sourceRange =
                                inputWorksheet.Cells[1, 1, lastRow, lastCol];

                            var destinationRange =
                                outputWorksheet.Cells[1, 1, lastRow, lastCol];

                            sourceRange.Copy(destinationRange);
                            outputWorksheet.InsertColumn(17, 1,16);
                            outputWorksheet.InsertRow(2, 3,3);
                            outputWorksheet.Cells[3, 3].Value = "State Bank of India ";
                            outputWorksheet.Cells[3, 5].Value = "00000010565627897";
                            outputWorksheet.Cells[3, 7].Value = "04327";
                            outputWorksheet.Cells[3, 6].Value = "#";
                            outputWorksheet.Cells[3, 8].Value = "#";
                            outputWorksheet.Cells[3, 10].Value = "#";
                            outputWorksheet.Cells[3, 12].Value = "##";
                            outputWorksheet.Cells[3, 14].Value = "#";
                            outputWorksheet.Cells[3, 16].Value = "#";
                            outputWorksheet.Cells[3, 9].Value = outputWorksheet.Cells[5, 9].Text;
                            outputWorksheet.Cells[3, 13].Value = outputWorksheet.Cells[5, 13].Text;
                            outputWorksheet.Cells[3, 15].Value = outputWorksheet.Cells[5, 15].Text;
                            outputWorksheet.Cells[1, 17].Value = "Final Outcome Result";
                            lastRow = outputWorksheet.Dimension.End.Row;
                            outputWorksheet.Cells[3, 11].Formula = $"SUM(K5:K{lastRow})";
                            // Auto Fit
                            outputWorksheet.Cells[3, 17].Formula = "=+E3&F3&G3&H3&I3&J3&K3&L3&M3&N3&O3&P3";
                            for (int i=5;i<=lastRow;i++)
                            {
                                outputWorksheet.Cells[i, 17].Formula = $"=+E{i}&F{i}&G{i}&H{i}&I{i}&J{i}&K{i}&L{i}&M{i}&N{i}&O{i}&P{i}";
                            }
                            outputWorksheet.Cells[
                                outputWorksheet.Dimension.Address].AutoFitColumns();
                        }
                    }

                    // ==============================
                    // COPY OTHER BANK FILE FIRST SHEET
                    // ==============================
                    using (var otherPackage = new ExcelPackage(new FileInfo(otherbankFile)))
                    {
                        var inputWorksheet = otherPackage.Workbook.Worksheets.FirstOrDefault();

                        if (inputWorksheet != null && inputWorksheet.Dimension != null)
                        {
                            int lastRow = inputWorksheet.Dimension.End.Row;
                            int lastCol = inputWorksheet.Dimension.End.Column;

                            // Create Sheet
                            var outputWorksheet =
                                outputPackage.Workbook.Worksheets.Add("Other Bank");

                            // Copy Data
                            var sourceRange =
                                inputWorksheet.Cells[1, 1, lastRow, lastCol];

                            var destinationRange =
                                outputWorksheet.Cells[1, 1, lastRow, lastCol];

                            sourceRange.Copy(destinationRange);
                            outputWorksheet.InsertColumn(19, 1,18);
                            outputWorksheet.InsertRow(2, 3, 3);
                            outputWorksheet.Cells[3, 4].Value = "State Bank of India ";
                            outputWorksheet.Cells[3, 6].Value = "00000010565627897";
                            outputWorksheet.Cells[3, 8].Value = "#";
                            outputWorksheet.Cells[3, 5].Value = "#";
                            outputWorksheet.Cells[3, 7].Value = "#";
                            outputWorksheet.Cells[3, 8].Value = "04327";
                            outputWorksheet.Cells[3, 9].Value = "#";
                            outputWorksheet.Cells[3, 11].Value = "#";
                            outputWorksheet.Cells[3, 13].Value = "##";
                            outputWorksheet.Cells[3, 15].Value = "#";
                            outputWorksheet.Cells[3, 17].Value = "#";
                            outputWorksheet.Cells[3, 10].Value = outputWorksheet.Cells[5, 10].Text;
                            outputWorksheet.Cells[3, 14].Value = outputWorksheet.Cells[5, 14].Text;
                            outputWorksheet.Cells[3, 16].Value = outputWorksheet.Cells[5, 16].Text;
                            outputWorksheet.Cells[3, 18].Value = outputWorksheet.Cells[5, 18].Text;
                            lastRow = outputWorksheet.Dimension.End.Row;
                            outputWorksheet.Cells[3, 12].Formula = $"SUM(L5:L{lastRow})";
                            outputWorksheet.Cells[1, 19].Value = "Payment Sheet";
                            outputWorksheet.Cells[3, 19].Formula = "=F3&G3&H3&I3&J3&K3&L3&M3&N3&O3&D3&Q3&R3";
                            for (int i = 5; i <= lastRow; i++)
                            {
                                outputWorksheet.Cells[i, 19].Formula = $"=F{i}&G{i}&H{i}&I{i}&J{i}&K{i}&L{i}&M{i}&N{i}&O{i}&D{i}&Q{i}&R{i}";
                            }
                            
                            outputWorksheet.Cells[
                                outputWorksheet.Dimension.Address].AutoFitColumns();
                        }
                    }

                    // =========================
                    // SAVE OUTPUT FILE
                    // =========================
                    string newFileName = Path.Combine(
                        outputFilePath,
                        "GH Induction Automated Bank File.xlsx");

                    FileInfo newFileInfo = new FileInfo(newFileName);

                    outputPackage.SaveAs(newFileInfo);

                    Console.WriteLine("GH Induction Automated Bank File created successfully.");
                    Console.WriteLine("Saved Path : " + newFileName);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error : " + ex.Message);
            }
        }
    }
}