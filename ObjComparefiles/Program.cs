using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using OfficeOpenXml;

class Program
{
    static string CalculateSHA256(string filePath)
    {
        using (var sha256 = SHA256.Create())
        {
            using (var stream = File.OpenRead(filePath))
            {
                byte[] hashBytes = sha256.ComputeHash(stream);
                string hashValue = BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
                return hashValue;
            }
        }
    }

    static List<string[]> ListFilesWithChecksum(string directory)
    {
        List<string[]> filesChecksums = new List<string[]>();
        int i = 0;
        foreach (string filePath in Directory.GetFiles(directory, "*", SearchOption.AllDirectories))
        {
            i++;
            string checksum = CalculateSHA256(filePath);
            filesChecksums.Add(new string[] { i.ToString(), filePath, checksum });
        }

        return filesChecksums;
    }

    static void SaveResultToExcel(List<string[]> result, string filePath, string directoryPath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (ExcelPackage package = new ExcelPackage())
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Comparison Result");

            // Set column headers
            worksheet.Cells[1, 1].Value = "Path : " + directoryPath;
            worksheet.Cells[2, 1].Value = "Num";
            worksheet.Cells[2, 2].Value = "Filename";
            worksheet.Cells[2, 3].Value = "Checksum";

            // Populate data rows
            for (int i = 0; i < result.Count; i++)
            {
                worksheet.Cells[i + 3, 1].Value = result[i][0];
                worksheet.Cells[i + 3, 2].Value = result[i][1];
                worksheet.Cells[i + 3, 3].Value = result[i][2];
            }

            package.SaveAs(new FileInfo(directoryPath + @"\" + filePath));
        }
    }

    static void Main()
    {
        Console.Write("Enter the directory path to checksum files: ");
        string directoryPath = Console.ReadLine();

        List<string[]> currentResult = ListFilesWithChecksum(directoryPath);


        SaveResultToExcel(currentResult, "resultchecksum256.xlsx", directoryPath);
    }
}
