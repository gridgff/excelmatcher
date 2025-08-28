using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeOpenXml;

namespace ExcelMatcher
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Excel Matcher Program");
            Console.WriteLine("=====================");
            
            try
            {
                
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                
                Console.Write("Enter the path to the Excel file: ");
                string inputFilePath = Console.ReadLine();

                if (string.IsNullOrEmpty(inputFilePath) || !File.Exists(inputFilePath))
                {
                    Console.WriteLine("File not found or invalid path.");
                    return;
                }

                
                ProcessExcelFile(inputFilePath);

                Console.WriteLine("Processing completed successfully!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
            }

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        static void ProcessExcelFile(string filePath)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                
                var sheet1 = package.Workbook.Worksheets["Лист_1"];
                var sheet2 = package.Workbook.Worksheets["Лист1"];

                if (sheet1 == null || sheet2 == null)
                {
                    throw new Exception("Required worksheets not found. Make sure both 'Лист_1' and 'Лист1' exist.");
                }

                
                var list1Data = ReadSheet1Data(sheet1);
                var list2Data = ReadSheet2Data(sheet2);

                
                var matchedData = MatchData(list1Data, list2Data);

                
                CreateOutputFile(matchedData, filePath);
            }
        }

        static List<Sheet1Record> ReadSheet1Data(ExcelWorksheet sheet)
        {
            var data = new List<Sheet1Record>();
            int rowCount = sheet.Dimension.Rows;
            int colCount = sheet.Dimension.Columns;

            
            int fioCol = -1, emailCol = -1;

            for (int col = 1; col <= colCount; col++)
            {
                string header = sheet.Cells[1, col].Text.Trim();
                if (header == "ФИО") fioCol = col;
                if (header == "Почта") emailCol = col;
            }

            if (fioCol == -1 || emailCol == -1)
            {
                throw new Exception("Required columns not found in Лист_1. Need 'ФИО' and 'Почта'.");
            }

            
            for (int row = 2; row <= rowCount; row++)
            {
                string fio = sheet.Cells[row, fioCol].Text.Trim();
                string email = sheet.Cells[row, emailCol].Text.Trim();

                if (!string.IsNullOrEmpty(fio) && !string.IsNullOrEmpty(email))
                {
                    data.Add(new Sheet1Record
                    {
                        FIO = fio,
                        Email = email,
                        Username = ExtractUsernameFromEmail(email)
                    });
                }
            }

            return data;
        }

        static List<Sheet2Record> ReadSheet2Data(ExcelWorksheet sheet)
        {
            var data = new List<Sheet2Record>();
            int rowCount = sheet.Dimension.Rows;
            int colCount = sheet.Dimension.Columns;

            
            int networkCodeCol = -1, accountCol = -1, ipcountCol = -1;

            for (int col = 1; col <= colCount; col++)
            {
                string header = sheet.Cells[1, col].Text.Trim();
                if (header == "Сетевой код") networkCodeCol = col;
                if (header == "Учетная запись") accountCol = col;
                if (header == "IP") ipcountCol = col;
            }

            if (networkCodeCol == -1 || accountCol == -1 || ipcountCol == -1)
            {
                throw new Exception("Required columns not found in Лист1. Need 'Сетевой код' and 'Учетная запись' and 'IP'.");
            }

            
            for (int row = 2; row <= rowCount; row++)
            {
                string networkCode = sheet.Cells[row, networkCodeCol].Text.Trim();
                string account = sheet.Cells[row, accountCol].Text.Trim();
                string ip = sheet.Cells[row, ipcountCol].Text.Trim();
                if (!string.IsNullOrEmpty(networkCode) && !string.IsNullOrEmpty(account) && !string.IsNullOrEmpty(ip))
                {
                    data.Add(new Sheet2Record
                    {
                        NetworkCode = networkCode,
                        Account = account,
                        Ip = ip,
                        Username = ExtractUsernameFromAccount(account)
                    });
                }
            }

            return data;
        }

        static string ExtractUsernameFromEmail(string email)
        {
            
            int atIndex = email.IndexOf('@');
            if (atIndex > 0)
            {
                return email.Substring(0, atIndex).ToLower();
            }
            return email.ToLower();
        }

        static string ExtractUsernameFromAccount(string account)
        {
            
            int backslashIndex = account.IndexOf('\\');
            if (backslashIndex > 0 && backslashIndex < account.Length - 1)
            {
                return account.Substring(backslashIndex + 1).ToLower();
            }
            return account.ToLower();
        }

        static List<MatchedRecord> MatchData(List<Sheet1Record> list1, List<Sheet2Record> list2)
        {
            var matchedData = new List<MatchedRecord>();

            foreach (var record1 in list1)
            {
                var matchingRecord = list2.FirstOrDefault(r => 
                    r.Username == record1.Username || 
                    r.Account.ToLower().Contains(record1.Username) ||
                    record1.Email.ToLower().Contains(r.Username));

                if (matchingRecord != null)
                {
                    matchedData.Add(new MatchedRecord
                    {
                        FIO = record1.FIO,
                        Email = record1.Email,
                        NetworkCode = matchingRecord.NetworkCode,
                        Ip = matchingRecord.Ip
                    });
                }
            }

            return matchedData;
        }

        static void CreateOutputFile(List<MatchedRecord> data, string originalFilePath)
        {
            string outputFilePath = Path.Combine(
                Path.GetDirectoryName(originalFilePath),
                $"Matched_Results_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

            using (var outputPackage = new ExcelPackage())
            {
                var worksheet = outputPackage.Workbook.Worksheets.Add("Matched Results");

                // Add headers
                worksheet.Cells[1, 1].Value = "ФИО";
                worksheet.Cells[1, 2].Value = "Почта";
                worksheet.Cells[1, 3].Value = "Имя компьютера";
                worksheet.Cells[1, 4].Value = "IP";

                // Style headers
                using (var range = worksheet.Cells[1, 1, 1, 3])
                {
                    range.Style.Font.Bold = true;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                }

                // Add data
                for (int i = 0; i < data.Count; i++)
                {
                    worksheet.Cells[i + 2, 1].Value = data[i].FIO;
                    worksheet.Cells[i + 2, 2].Value = data[i].Email;
                    worksheet.Cells[i + 2, 3].Value = data[i].NetworkCode;
                    worksheet.Cells[i + 2, 4].Value = data[i].Ip;
                }

                // Auto-fit columns
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                // Save the file
                outputPackage.SaveAs(new FileInfo(outputFilePath));
            }

            Console.WriteLine($"Output file created: {outputFilePath}");
            Console.WriteLine($"Total matched records: {data.Count}");
        }
    }

    // Data classes
    class Sheet1Record
    {
        public string FIO { get; set; }
        public string Email { get; set; }
        public string Username { get; set; }
    }

    class Sheet2Record
    {
        public string NetworkCode { get; set; }
        public string Account { get; set; }

        public string Username { get; set; }
        public string Ip { get; set; }
    }

    class MatchedRecord
    {
        public string FIO { get; set; }
        public string Email { get; set; }
        public string NetworkCode { get; set; }
        public string Ip { get; set; }
    }
}